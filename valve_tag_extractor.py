#!/usr/bin/env python3
"""P&ID Valve Tag Extractor.

Scans every PDF in a folder (recursively) and extracts valve tags such as
GLV-0022, GV-0220, SPV 3221B, SPI 001 ... together with the file/page they
came from.

Unlike plain text scraping, PDF *annotations* (mark-ups) are evaluated:

* A tag hidden by an opaque annotation (white-out box, filled polygon,
  redaction, stamp, strike-out ...) is reported as DELETED.
* A tag partly hidden by an annotation that carries new text (FreeText /
  text box, or a white-out box with a text box on top) is reported as
  CHANGED: the new tag number is reported as the tag, and the original
  number that is still in the page content is reported as "Original Tag".
* A valve tag written only in an annotation (not over an existing tag) is
  reported as ADDED.

Usage:
    python valve_tag_extractor.py [FOLDER] [-o valve_tags.xlsx]
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

try:
    import pymupdf  # PyMuPDF >= 1.24
except ImportError:  # pragma: no cover - older PyMuPDF
    try:
        import fitz as pymupdf  # type: ignore
    except ImportError:
        sys.exit("PyMuPDF is required:  pip install pymupdf openpyxl")


DEFAULT_PREFIXES = ["BFV", "BV", "CV", "GLV", "GV", "SPV", "SPI", "GVL", "GA"]

# Status values
ACTIVE = "ACTIVE"
CHANGED = "CHANGED"
DELETED = "DELETED"
ADDED = "ADDED"
CHECK = "CHECK"

# Red box mark-up written to the copies in the mark-up folder
MARKUP_DIR = "_tagged"
MARKUP_AUTHOR = "Valve Tag Extractor"
MARKUP_PAD = 1.5

# Fraction of a tag's characters that must be hidden to call it deleted.
DELETE_COVER_RATIO = 0.6

HYPHENS = "‐‑‒–—−﹣－"
HYPHEN_TRANS = str.maketrans({h: "-" for h in HYPHENS})

# Annotation types that may hide content when they have a fill / opacity.
FILLABLE_TYPES = {"Square", "Circle", "Polygon"}
ALWAYS_COVER_TYPES = {"Redact", "Stamp"}
STRIKE_TYPES = {"StrikeOut"}
TEXT_ANNOT_TYPES = {"FreeText"}


def build_patterns(prefixes):
    # Longest prefix first so GVL / GLV are not matched as GV.
    alt = "|".join(sorted((re.escape(p) for p in prefixes), key=len, reverse=True))
    number = r"\d{2,6}[A-Z]?"
    full = re.compile(rf"(?<![A-Z0-9])({alt})\s*-?\s*({number})(?![A-Z0-9])")
    prefix_only = re.compile(rf"^\s*({alt})\s*-?\s*$")
    number_only = re.compile(rf"^\s*-?\s*({number})\s*$")
    return full, prefix_only, number_only


def normalize_text(s: str) -> str:
    return s.translate(HYPHEN_TRANS).upper()


def make_tag(prefix: str, number: str) -> str:
    return f"{prefix.upper()}-{number.upper()}"


@dataclass
class Char:
    c: str
    rect: pymupdf.Rect

    @property
    def center(self):
        return pymupdf.Point((self.rect.x0 + self.rect.x1) / 2,
                             (self.rect.y0 + self.rect.y1) / 2)


@dataclass
class Line:
    chars: list  # list[Char]
    rect: pymupdf.Rect

    @property
    def text(self):
        return "".join(ch.c for ch in self.chars)

    @property
    def height(self):
        return max(1.0, min(self.rect.width, self.rect.height))


@dataclass
class Annotation:
    kind: str
    rect: pymupdf.Rect
    author: str
    content: str
    text: str            # visible text (FreeText) normalised
    covers: bool          # hides what is beneath it
    strikes: bool
    quads: list = field(default_factory=list)
    used: bool = False
    has_image: bool = False   # appearance contains a picture (pasted image / image stamp)

    def describe(self):
        parts = [self.kind]
        if self.text:
            parts.append(f'text="{self.text}"')
        if self.content and self.content != self.text:
            parts.append(f'comment="{self.content}"')
        if self.author:
            parts.append(f"by {self.author}")
        return " ".join(parts)


@dataclass
class TagHit:
    file: str
    page: int
    tag: str
    status: str
    original_tag: str
    rect: pymupdf.Rect
    note: str = ""
    annotations: str = ""
    source: str = "TEXT"
    score: float | None = None


# --------------------------------------------------------------------------
# Page parsing
# --------------------------------------------------------------------------

def page_lines(page) -> list:
    """Text lines from the page *content* (annotations are not included)."""
    lines = []
    raw = page.get_text("rawdict", flags=pymupdf.TEXT_PRESERVE_WHITESPACE)
    for block in raw.get("blocks", []):
        for line in block.get("lines", []):
            chars = []
            for span in line.get("spans", []):
                for ch in span.get("chars", []):
                    chars.append(Char(normalize_text(ch["c"]), pymupdf.Rect(ch["bbox"])))
            if chars and "".join(c.c for c in chars).strip():
                lines.append(Line(chars, pymupdf.Rect(line["bbox"])))
    return lines


def _fill_is_visible(colors) -> bool:
    fill = colors.get("fill") if colors else None
    return bool(fill)


def annot_visible_text(annot) -> str:
    text = ""
    try:
        text = annot.get_text("text") or ""
    except Exception:
        text = ""
    if not text.strip():
        text = annot.info.get("content", "") or ""
    return " ".join(normalize_text(text).split())


def page_annotations(page, warnings=None) -> list:
    result = []
    try:
        annot_list = list(page.annots() or [])
    except Exception as exc:
        if warnings is not None:
            warnings.append(f"p{page.number + 1}: cannot read annotations ({exc})")
        return result
    for annot in annot_list:
        try:
            item = _read_annotation(annot)
        except Exception as exc:  # one broken mark-up must not fail the file
            if warnings is not None:
                warnings.append(f"p{page.number + 1}: skipped annotation ({exc})")
            continue
        if item is not None:
            result.append(item)
    return result


def _read_annotation(annot):
    kind = annot.type[1]
    if kind in ("Popup", "Link", "Widget"):
        return None
    try:
        if annot.flags & pymupdf.PDF_ANNOT_IS_HIDDEN:
            return None
    except Exception:
        pass
    opacity = annot.opacity if annot.opacity is not None and annot.opacity >= 0 else 1.0
    info = annot.info or {}
    content = " ".join((info.get("content") or "").split())
    text = annot_visible_text(annot) if kind in TEXT_ANNOT_TYPES else ""
    covers = False
    if opacity >= 0.5:
        if kind in ALWAYS_COVER_TYPES:
            covers = True
        elif kind in FILLABLE_TYPES and _fill_is_visible(annot.colors):
            covers = True
        elif kind in TEXT_ANNOT_TYPES:
            covers = True  # a text box written over a tag replaces it
    quads = []
    if kind in STRIKE_TYPES:
        verts = annot.vertices or []
        for i in range(0, len(verts) - 3, 4):
            quads.append(pymupdf.Quad(*verts[i:i + 4]).rect)
        if not quads:
            quads = [annot.rect]
    has_image = False
    try:
        ap = annot.parent.parent.xref_get_key(annot.xref, "AP/N")
        if ap[0] == "xref":
            has_image = _xobject_has_image(annot.parent.parent, int(ap[1].split()[0]))
    except Exception:
        pass
    if has_image and opacity >= 0.5:
        covers = True  # an opaque picture hides the drawing below it
    return Annotation(
        has_image=has_image,
        kind=kind,
        rect=pymupdf.Rect(annot.rect),
        author=info.get("title", "") or "",
        content=content,
        text=text,
        covers=covers,
        strikes=kind in STRIKE_TYPES,
        quads=quads,
    )


def _xobject_has_image(doc, xref, depth=0) -> bool:
    """True if the form XObject (annotation appearance) draws an image."""
    if depth > 5:
        return False
    if doc.xref_get_key(xref, "Subtype")[1] == "/Image":
        return True
    typ, val = doc.xref_get_key(xref, "Resources/XObject")
    if typ == "xref":
        val = doc.xref_object(int(val.split()[0]), compressed=True)
    for ref in re.findall(r"(\d+) 0 R", val if typ in ("dict", "xref") else ""):
        if _xobject_has_image(doc, int(ref), depth + 1):
            return True
    return False


# --------------------------------------------------------------------------
# OCR of pictures pasted into the drawing
# --------------------------------------------------------------------------

OCR_DPI = 300            # render resolution for pictures
OCR_MIN_SCORE = 0.5      # discard OCR text below this confidence
OCR_TILE_PX = 1600       # large pictures are OCR'd in tiles of this size
MIN_IMAGE_PT = 20        # ignore icons / tiny pictures (points)
OCR_FIX = str.maketrans({"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1", "|": "1"})


class OcrEngine:
    """Lazy wrapper around RapidOCR (pip install rapidocr_onnxruntime)."""

    def __init__(self):
        self._engine = None
        self.error = None

    def available(self) -> bool:
        if self._engine is None and self.error is None:
            try:
                from rapidocr_onnxruntime import RapidOCR
                self._engine = RapidOCR()
                self._api = 1
            except ImportError:
                try:
                    from rapidocr import RapidOCR
                    self._engine = RapidOCR()
                    self._api = 2
                except ImportError:
                    self.error = ("OCR not installed (pip install rapidocr_onnxruntime) - "
                                  "tags inside pasted pictures are skipped")
        return self._engine is not None

    def __call__(self, img):
        """Return [(4 corner points in pixels, text, score)]."""
        if self._api == 1:
            result, _ = self._engine(img)
            return [(r[0], r[1], float(r[2])) for r in (result or [])]
        out = self._engine(img)
        if out.boxes is None:
            return []
        return [(b.tolist(), t, float(sc)) for b, t, sc in zip(out.boxes, out.txts, out.scores)]


def _pixmap_to_array(pix):
    import numpy as np
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    return arr[:, :, :3].copy() if pix.n >= 3 else np.repeat(arr[:, :, :1], 3, axis=2)


def fix_ocr_text(text, prefix_alt):
    """Uppercase and repair typical OCR confusions in the number part of a tag."""
    text = normalize_text(text).replace("_", "-").replace("~", "-")

    def repl(m):
        return m.group(1) + m.group(2).replace("=", "-") + m.group(3).translate(OCR_FIX)

    return re.sub(rf"(?<![A-Z0-9])({prefix_alt})(\s*[-.=]?\s*)([0-9OQDIL|]{{2,6}})(?=[A-Z]?(?![A-Z0-9]))",
                  repl, text)


def ocr_region(page, rect, engine, prefix_alt, annots_shown=False):
    """OCR one area of the page; returns [(Line in page coordinates, score)]."""
    lines = []
    tile = OCR_TILE_PX * 72.0 / OCR_DPI
    overlap = 40.0
    seen = []
    y = rect.y0
    while y < rect.y1:
        x = rect.x0
        while x < rect.x1:
            clip = pymupdf.Rect(x, y, min(x + tile, rect.x1), min(y + tile, rect.y1))
            if clip.width > 2 and clip.height > 2:
                pix = page.get_pixmap(clip=clip, dpi=OCR_DPI, annots=annots_shown, alpha=False)
                sx = clip.width / pix.width
                sy = clip.height / pix.height
                for box, text, score in engine(_pixmap_to_array(pix)):
                    if score < OCR_MIN_SCORE or not text.strip():
                        continue
                    xs = [clip.x0 + pt[0] * sx for pt in box]
                    ys = [clip.y0 + pt[1] * sy for pt in box]
                    r = pymupdf.Rect(min(xs), min(ys), max(xs), max(ys))
                    text = fix_ocr_text(text, prefix_alt)
                    if any(t == text and abs(r.x0 - o.x0) < 3 and abs(r.y0 - o.y0) < 3 for t, o in seen):
                        continue  # same text found again in a tile overlap
                    seen.append((text, r))
                    lines.append((_ocr_line(text, r), score))
            x += tile - overlap
        y += tile - overlap
    return lines


def _ocr_line(text, r):
    """Spread the characters of an OCR result evenly over its box."""
    n = max(1, len(text))
    chars = []
    vertical = r.height > r.width * 1.5 and n > 2
    for i, c in enumerate(text):
        if vertical:
            h = r.height / n
            cr = pymupdf.Rect(r.x0, r.y1 - (i + 1) * h, r.x1, r.y1 - i * h)  # bottom-to-top text
        else:
            w = r.width / n
            cr = pymupdf.Rect(r.x0 + i * w, r.y0, r.x0 + (i + 1) * w, r.y1)
        chars.append(Char(c, cr))
    return Line(chars, pymupdf.Rect(r))


def content_pictures(clean_page):
    """Pictures in the page content: [(rect, seqno, opaque)]."""
    pics = []
    for info in clean_page.get_image_info():
        r = pymupdf.Rect(info["bbox"]) & clean_page.mediabox
        if r.is_empty or r.width < MIN_IMAGE_PT or r.height < MIN_IMAGE_PT:
            continue
        pics.append((r, info.get("number", -1), not info.get("has-mask", False)))
    return pics


def hidden_by_pictures(clean_page, pics):
    """Text drawn *before* an opaque picture inside its area is not visible.

    Returned as pseudo annotations so evaluate_tag() treats them like a
    white-out box.
    """
    covers = []
    opaque = [(r, seq) for r, seq, op in pics if op and seq >= 0]
    if not opaque:
        return covers
    for span in clean_page.get_texttrace():
        sr = pymupdf.Rect(span["bbox"])
        c = pymupdf.Point((sr.x0 + sr.x1) / 2, (sr.y0 + sr.y1) / 2)
        for r, seq in opaque:
            if span["seqno"] < seq and r.contains(c):
                covers.append(Annotation("Pasted picture", sr, "", "", "", True, False))
                break
    return covers


def find_tags_in_lines(lines, full_re, prefix_re, number_re):
    """Return list of (prefix, number, chars) found in the page content."""
    found = []
    used_lines = set()

    # 1) Tags written in one line: "GLV-0022", "SPV 3221B"
    for li, line in enumerate(lines):
        text = line.text
        for m in full_re.finditer(text):
            chars = line.chars[m.start():m.end()]
            found.append((m.group(1), m.group(2), chars))
            used_lines.add(li)

    # 2) Prefix and number in separate lines (SPV / SPI boxes, "GV-" + "0220")
    prefixes = [(i, l) for i, l in enumerate(lines)
                if i not in used_lines and prefix_re.match(l.text)]
    numbers = [(i, l) for i, l in enumerate(lines)
               if i not in used_lines and number_re.match(l.text)]
    taken = set()
    pairs = []
    for pi, pl in prefixes:
        h = pl.height
        for ni, nl in numbers:
            dx = nl.rect.x0 - pl.rect.x1          # gap to the right
            dy = nl.rect.y0 - pl.rect.y1          # gap below
            horiz_overlap = min(nl.rect.x1, pl.rect.x1) - max(nl.rect.x0, pl.rect.x0)
            vert_overlap = min(nl.rect.y1, pl.rect.y1) - max(nl.rect.y0, pl.rect.y0)
            ok = False
            # right of the prefix on the same text line
            if vert_overlap > 0.5 * h and -0.5 * h <= dx <= 2.0 * h:
                ok = True
            # below the prefix (stacked / instrument-box style)
            elif -0.5 * h <= dy <= 3.0 * h and (horiz_overlap > -4.0 * h):
                ok = True
            if ok:
                dist = math.hypot(max(dx, 0), max(dy, 0))
                pairs.append((dist, pi, ni))
    for dist, pi, ni in sorted(pairs):
        if pi in taken or ni in taken:
            continue
        taken.update((pi, ni))
        pl, nl = lines[pi], lines[ni]
        prefix = prefix_re.match(pl.text).group(1)
        number = number_re.match(nl.text).group(1)
        chars = [c for c in pl.chars if c.c.strip()] + [c for c in nl.chars if c.c.strip()]
        found.append((prefix, number, chars))
    return found


# --------------------------------------------------------------------------
# Annotation evaluation
# --------------------------------------------------------------------------

def _inside(pt, rect, tol=0.5):
    return (rect.x0 - tol <= pt.x <= rect.x1 + tol) and (rect.y0 - tol <= pt.y <= rect.y1 + tol)


def _union(rects):
    r = pymupdf.Rect(rects[0])
    for x in rects[1:]:
        r |= x
    return r


def evaluate_tag(prefix, number, chars, annots, full_re, number_re):
    """Decide ACTIVE / CHANGED / DELETED for one tag found in the content."""
    original = make_tag(prefix, number)
    visible = [c for c in chars if c.c.strip()]
    rect = _union([c.rect for c in visible])

    covering = [a for a in annots if a.covers and a.rect.intersects(rect)]
    striking = [a for a in annots if a.strikes and any(q.intersects(rect) for q in a.quads)]

    hidden = [any(_inside(c.center, a.rect) for a in covering) for c in visible]
    n_hidden = sum(hidden)
    ratio = n_hidden / len(visible) if visible else 0.0

    # Text annotations sitting on (or overlapping) the tag
    texts = [a for a in covering if a.kind in TEXT_ANNOT_TYPES and a.text]

    involved = covering + striking
    ann_desc = "; ".join(a.describe() for a in involved)

    if texts and n_hidden:
        new_text = " ".join(a.text for a in texts)
        for a in texts:
            a.used = True
        m = full_re.search(new_text)
        if m:
            new_tag = make_tag(m.group(1), m.group(2))
        else:
            # Replace the hidden part of the original tag text with the new text.
            tag_str = "".join(c.c for c in visible)
            idx = [i for i, h in enumerate(hidden) if h]
            start, end = idx[0], idx[-1] + 1
            spliced = normalize_text(tag_str[:start] + new_text.replace(" ", "") + tag_str[end:])
            m2 = full_re.search(spliced)
            if m2:
                new_tag = make_tag(m2.group(1), m2.group(2))
            else:
                m3 = number_re.match(new_text)
                new_tag = make_tag(prefix, m3.group(1)) if m3 else None
        if new_tag and new_tag != original:
            return CHANGED, new_tag, rect, "Number changed by annotation", ann_desc
        if new_tag == original:
            return ACTIVE, original, rect, "Annotation re-writes the same tag", ann_desc
        return CHECK, original, rect, f'Annotation text over tag not understood: "{new_text}"', ann_desc

    if striking:
        return DELETED, original, rect, "Struck out by annotation", ann_desc

    if n_hidden and ratio >= DELETE_COVER_RATIO:
        return DELETED, original, rect, f"Hidden by annotation ({n_hidden}/{len(visible)} chars)", ann_desc

    if n_hidden:
        return CHECK, original, rect, f"Partly hidden by annotation ({n_hidden}/{len(visible)} chars), no new text found", ann_desc

    # Non-covering markup touching the tag (revision clouds, outlines, ...) - informational
    nearby = [a for a in annots if not a.covers and not a.strikes and a.rect.intersects(rect)]
    if nearby:
        ann_desc = "; ".join(a.describe() for a in nearby)
    return ACTIVE, original, rect, "", ann_desc


def ocr_pictures(page, clean_page, pics, pic_annots, annots, ocr, page_hits,
                 rel, pno, full_re, prefix_re, number_re, prefix_alt, warnings):
    """Add the tags read by OCR from pasted pictures to page_hits."""
    if not ocr.available():
        if warnings is not None:
            warnings.append(f"p{pno}: {len(pics) + len(pic_annots)} picture(s) not read - {ocr.error}")
        return page_hits
    # Render unrotated (in memory only) so OCR boxes map straight to page coordinates.
    clean_page.set_rotation(0)
    page.set_rotation(0)
    regions = [(r, clean_page, False, None, "PICTURE (OCR)") for r, _, _ in pics]
    regions += [(a.rect, page, True, a, "ANNOT PICTURE (OCR)") for a in pic_annots]
    for rect, src_page, shown, src_annot, source in regions:
        try:
            ocr_lines = ocr_region(src_page, rect, ocr, prefix_alt, shown)
        except Exception as exc:
            if warnings is not None:
                warnings.append(f"p{pno}: OCR failed ({exc})")
            continue
        score_of = {id(c): sc for ln, sc in ocr_lines for c in ln.chars}
        others = [a for a in annots if a is not src_annot]
        for prefix, number, chars in find_tags_in_lines([ln for ln, _ in ocr_lines],
                                                        full_re, prefix_re, number_re):
            status, tag, trect, note, ann = evaluate_tag(prefix, number, chars, others, full_re, number_re)
            score = min(score_of.get(id(c), 1.0) for c in chars)
            if src_annot is not None:
                ann = "; ".join(x for x in (src_annot.describe(), ann) if x)
            page_hits.append(TagHit(rel, pno, tag, status, make_tag(prefix, number), trect,
                                    note or "Read from pasted picture", ann, source=source, score=score))
    return merge_ocr_duplicates(page_hits)


def merge_ocr_duplicates(hits):
    """Drop OCR hits that repeat a text tag at the same place.

    A picture rendered for OCR also shows drawing text lying on top of it, and
    a tag pasted again inside a picture hides the original text below it.  If
    both name the same tag at the same place, keep the visible one.
    """
    text_hits = [h for h in hits if h.source == "TEXT"]
    kept_ocr = []
    drop = set()
    for o in (h for h in hits if h.source != "TEXT"):
        area = pymupdf.Rect(o.rect) + (-3, -3, 3, 3)
        same = [t for t in text_hits if t.tag == o.tag and area.intersects(t.rect)]
        if same and any(t.status != DELETED for t in same):
            continue  # visible drawing text, already reported
        drop.update(id(t) for t in same)  # picture re-shows a tag hidden below it
        if not any(k.tag == o.tag and area.intersects(k.rect) for k in kept_ocr):
            kept_ocr.append(o)
    return [h for h in text_hits if id(h) not in drop] + kept_ocr


def process_pdf(path: Path, root: Path, full_re, prefix_re, number_re, warnings=None, ocr=None):
    hits = []
    prefix_alt = full_re.pattern.split("(", 2)[2].split(")", 1)[0]
    rel = str(path.relative_to(root)) if path.is_relative_to(root) else str(path)
    with pymupdf.open(path) as doc, pymupdf.open(path) as clean:
        for pno, page in enumerate(doc, start=1):
            annots = page_annotations(page, warnings)
            # MuPDF text extraction also returns annotation text, so read the
            # original drawing text from an in-memory copy without annotations.
            # The page's /Annots list is dropped as a whole: deleting comments
            # one by one also deletes their replies (IRT) and review states,
            # which then breaks the iteration ("Annot is not bound to a page").
            clean.xref_set_key(clean.page_xref(pno - 1), "Annots", "null")
            clean_page = clean[pno - 1]
            lines = page_lines(clean_page)

            # Pictures pasted into the drawing (page content or image stamps).
            pics = content_pictures(clean_page)
            pic_annots = [a for a in annots if a.has_image
                          and a.rect.width >= MIN_IMAGE_PT and a.rect.height >= MIN_IMAGE_PT]
            text_annots = annots + hidden_by_pictures(clean_page, pics) if pics else annots

            page_hits = []
            for prefix, number, chars in find_tags_in_lines(lines, full_re, prefix_re, number_re):
                status, tag, rect, note, ann = evaluate_tag(prefix, number, chars, text_annots, full_re, number_re)
                page_hits.append(TagHit(rel, pno, tag, status,
                                        make_tag(prefix, number), rect, note, ann))

            if ocr is not None and (pics or pic_annots):
                page_hits = ocr_pictures(page, clean_page, pics, pic_annots, annots, ocr, page_hits,
                                         rel, pno, full_re, prefix_re, number_re, prefix_alt, warnings)
            hits.extend(page_hits)
            # Tags that exist only inside annotations
            for a in annots:
                if a.used or not a.text:
                    continue
                for m in full_re.finditer(a.text):
                    hits.append(TagHit(rel, pno, make_tag(m.group(1), m.group(2)), ADDED, "",
                                       a.rect, "Tag written only in an annotation", a.describe(),
                                       source="ANNOT TEXT"))
    hits.sort(key=lambda h: (h.page, round(h.rect.y0), h.rect.x0))
    return hits


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------

COLUMNS = ["File", "Page", "Valve Tag", "Status", "Original Tag (text)",
           "Note", "Annotation", "Source", "OCR Score", "X", "Y"]


def hit_row(h: TagHit):
    return [h.file, h.page, h.tag, h.status,
            h.original_tag if h.status in (CHANGED,) else ("" if h.status == ADDED else h.original_tag),
            h.note, h.annotations, h.source, round(h.score, 2) if h.score is not None else "",
            round(h.rect.x0, 1), round(h.rect.y0, 1)]


def write_markup(src: Path, dst: Path, hits, warnings=None):
    """Save a copy of the PDF with a red box around every extracted tag.

    The box's comment (contents) is the extracted valve tag, so it can be
    reviewed / searched in the PDF viewer's comment list.
    """
    dst.parent.mkdir(parents=True, exist_ok=True)
    with pymupdf.open(src) as doc:
        for h in hits:
            try:
                page = doc[h.page - 1]
                rect = pymupdf.Rect(h.rect) + (-MARKUP_PAD, -MARKUP_PAD, MARKUP_PAD, MARKUP_PAD)
                annot = page.add_rect_annot(rect)
                annot.set_colors(stroke=(1, 0, 0))
                annot.set_border(width=1)
                annot.set_info(content=h.tag, title=MARKUP_AUTHOR, subject=h.status)
                annot.update()
            except Exception as exc:
                if warnings is not None:
                    warnings.append(f"p{h.page}: red box for {h.tag} failed ({exc})")
        doc.save(dst, garbage=1, deflate=True)


def write_csv(path: Path, hits, errors):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(COLUMNS)
        for h in hits:
            w.writerow(hit_row(h))
        for file, err in errors:
            w.writerow([file, "", "", "ERROR", "", err] + [""] * 5)


def write_xlsx(path: Path, hits, errors):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    fills = {
        CHANGED: PatternFill("solid", fgColor="FFF2CC"),
        DELETED: PatternFill("solid", fgColor="F8CBAD"),
        ADDED: PatternFill("solid", fgColor="C6EFCE"),
        CHECK: PatternFill("solid", fgColor="BDD7EE"),
    }
    wb = Workbook()
    ws = wb.active
    ws.title = "Valve Tags"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for h in hits:
        ws.append(hit_row(h))
        if h.status in fills:
            for c in ws[ws.max_row]:
                c.fill = fills[h.status]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, width in enumerate([40, 6, 16, 10, 18, 45, 60, 20, 10, 8, 8], start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Summary per file
    ss = wb.create_sheet("Summary")
    ss.append(["File", "Total", ACTIVE, CHANGED, DELETED, ADDED, CHECK])
    for c in ss[1]:
        c.font = Font(bold=True)
    per_file = {}
    for h in hits:
        d = per_file.setdefault(h.file, {k: 0 for k in (ACTIVE, CHANGED, DELETED, ADDED, CHECK)})
        d[h.status] += 1
    for file, d in per_file.items():
        ss.append([file, sum(d.values()), d[ACTIVE], d[CHANGED], d[DELETED], d[ADDED], d[CHECK]])
    ss.column_dimensions["A"].width = 50

    # Unique current tags (deleted ones excluded) with the files they appear in
    us = wb.create_sheet("Unique Tags")
    us.append(["Valve Tag", "Status(es)", "Count", "Files"])
    for c in us[1]:
        c.font = Font(bold=True)
    uniq = {}
    for h in hits:
        d = uniq.setdefault(h.tag, {"status": set(), "count": 0, "files": []})
        d["status"].add(h.status)
        d["count"] += 1
        loc = f"{h.file} (p{h.page})"
        if loc not in d["files"]:
            d["files"].append(loc)
    for tag in sorted(uniq):
        d = uniq[tag]
        us.append([tag, ", ".join(sorted(d["status"])), d["count"], " | ".join(d["files"])])
    us.column_dimensions["A"].width = 16
    us.column_dimensions["B"].width = 20
    us.column_dimensions["D"].width = 100

    if errors:
        es = wb.create_sheet("Errors")
        es.append(["File", "Error"])
        for e in errors:
            es.append(list(e))
    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Extract valve tags from P&ID PDFs (annotation aware).")
    ap.add_argument("folder", nargs="?", default=".", help="folder with PDF files (default: current folder)")
    ap.add_argument("-o", "--output", default="valve_tags.xlsx",
                    help="output file (.xlsx or .csv, default: valve_tags.xlsx)")
    ap.add_argument("--prefixes", default=",".join(DEFAULT_PREFIXES),
                    help="comma separated tag prefixes (default: %(default)s)")
    ap.add_argument("--markup-dir", default=MARKUP_DIR,
                    help="folder for PDF copies with red boxes around extracted tags (default: %(default)s)")
    ap.add_argument("--no-markup", action="store_true", help="do not write red-box PDF copies")
    ap.add_argument("--no-ocr", action="store_true",
                    help="do not OCR pictures pasted into the drawings")
    ap.add_argument("--no-recursive", action="store_true", help="do not search sub folders")
    args = ap.parse_args(argv)

    root = Path(args.folder).resolve()
    if not root.is_dir():
        sys.exit(f"Folder not found: {root}")
    prefixes = [p.strip().upper() for p in args.prefixes.split(",") if p.strip()]
    full_re, prefix_re, number_re = build_patterns(prefixes)

    pattern = "*.pdf" if args.no_recursive else "**/*.pdf"
    pdfs = sorted(p for p in root.glob(pattern) if p.suffix.lower() == ".pdf")
    if args.no_recursive:
        pdfs += sorted(p for p in root.glob("*.PDF") if p not in pdfs)
    else:
        pdfs += sorted(p for p in root.glob("**/*.PDF") if p not in pdfs)
    markup_root = Path(args.markup_dir)
    if not markup_root.is_absolute():
        markup_root = root / markup_root
    markup_root = markup_root.resolve()
    pdfs = [p for p in pdfs if not p.resolve().is_relative_to(markup_root)]
    if not pdfs:
        sys.exit(f"No PDF files in {root}")

    # MuPDF prints repairable PDF problems to the console; they are not fatal.
    try:
        pymupdf.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass

    ocr = None if args.no_ocr else OcrEngine()
    hits, errors = [], []
    for i, pdf in enumerate(pdfs, start=1):
        print(f"[{i}/{len(pdfs)}] {pdf.name}", end="", flush=True)
        warnings = []
        try:
            file_hits = process_pdf(pdf, root, full_re, prefix_re, number_re, warnings, ocr)
        except Exception as exc:  # keep going with the other files
            errors.append((pdf.name, str(exc)))
            print("  -> ERROR:", exc)
            continue
        hits.extend(file_hits)
        counts = {}
        for h in file_hits:
            counts[h.status] = counts.get(h.status, 0) + 1
        print("  ->", len(file_hits), "tags", counts if counts else "")
        if not args.no_markup:
            try:
                write_markup(pdf, markup_root / pdf.relative_to(root), file_hits, warnings)
            except Exception as exc:
                warnings.append(f"red-box PDF not saved ({exc})")
        for w in warnings:
            print("     WARNING:", w)
            errors.append((pdf.name, "WARNING: " + w))

    out = Path(args.output)
    if not out.is_absolute():
        out = root / out
    if out.suffix.lower() == ".csv":
        write_csv(out, hits, errors)
    else:
        write_xlsx(out, hits, errors)
    print(f"\n{len(hits)} tags from {len(pdfs)} PDF(s) -> {out}")
    if not args.no_markup:
        print(f"Red-box PDFs -> {markup_root}")


if __name__ == "__main__":
    main()
