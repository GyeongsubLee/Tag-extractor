# P&ID Valve Tag Extractor

폴더 안의 모든 P&ID PDF에서 Valve Tag(`BFV, BV, CV, GLV, GV, SPV, SPI, GVL, GA`)를 추출하고,
**어떤 파일 / 몇 페이지**에서 나왔는지 엑셀로 정리합니다.

단순 텍스트 추출이 아니라 **PDF 주석(Annotation)** 을 함께 분석합니다.

| Status | 의미 |
|---|---|
| `ACTIVE` | 정상 Tag |
| `DELETED` | 주석(흰 박스, 채워진 도형, Redact, Stamp, 취소선 등)으로 가려진/지워진 Tag |
| `CHANGED` | 주석(텍스트 박스)으로 번호가 바뀐 Tag. `Valve Tag` = 바뀐 번호, `Original Tag (text)` = 원래 도면 텍스트 |
| `ADDED` | 기존 Tag 위가 아닌 곳에 주석으로만 새로 써 넣은 Tag |
| `CHECK` | 일부만 가려졌는데 새 번호를 찾지 못한 경우 등 → 사람이 확인 필요 |

예) `GLV-0024`의 `0024` 부분만 흰 박스로 가리고 텍스트 박스로 `0030`을 쓴 경우 →
`Valve Tag = GLV-0030`, `Status = CHANGED`, `Original Tag = GLV-0024`

### 붙여넣은 사진(이미지) 속 Tag
- 도면에 캡처 이미지를 붙여넣은 경우(도면 안에 삽입 / 이미지 스탬프 주석 모두) 그 영역만 **OCR**로 읽어서 Tag를 추출합니다.
  이미지가 없는 페이지는 OCR을 하지 않으므로 속도는 기존과 같습니다.
- `Source` 열: `TEXT`(도면 텍스트), `PICTURE (OCR)`(도면에 삽입된 사진), `ANNOT PICTURE (OCR)`(이미지 주석), `ANNOT TEXT`(텍스트 주석)
- `OCR Score` 열: OCR 신뢰도(0~1). 낮은 값은 한번 확인해 주세요.
- 사진 밑에 가려진 원래 도면 Tag는 `DELETED`(Annotation = `Pasted picture`)로 표시됩니다. 같은 Tag가 사진에 다시 있으면 사진 쪽 하나만 남깁니다.
- OCR 끄기: `--no-ocr`

## 설치

```
pip install -r requirements.txt
```

## 실행

```
python valve_tag_extractor.py "D:\PID\PDF폴더"
```

- 폴더를 생략하면 현재 폴더에서 실행합니다. 하위 폴더까지 검색합니다 (`--no-recursive` 로 끌 수 있음).
- 결과는 해당 폴더에 `valve_tags.xlsx` 로 저장됩니다 (`-o result.csv` 처럼 CSV도 가능).
- 추출된 Tag마다 **빨간 박스(RED BOX) 주석**을 친 PDF 사본이 `_tagged` 폴더에 원본과 같은 폴더 구조로 저장됩니다.
  - 주석 설명(Comment) = 추출된 Valve Tag (예: `GLV-0030`), 주석 제목(Subject) = 상태(`ACTIVE`/`CHANGED`/`DELETED`…), 작성자 = `Valve Tag Extractor`
  - 원본 PDF는 수정하지 않습니다. `_tagged` 폴더는 다음 실행 시 검색에서 제외됩니다.
  - 폴더 변경: `--markup-dir 폴더명`, 끄기: `--no-markup`
- Windows에서는 `run.bat` 을 PDF 폴더에서 실행해도 됩니다.
- Prefix 변경: `--prefixes BFV,BV,CV,GLV,GV,SPV,SPI,GVL,GA,XV`

### 엑셀 시트
- **Valve Tags**: 모든 Tag (파일, 페이지, Tag, 상태, 원래 Tag, 비고, 관련 주석, 좌표). 상태별 색상 표시
- **Summary**: 파일별 상태 개수
- **Unique Tags**: Tag별 등장 파일 목록
- **Errors**: 읽지 못한 파일

## 동작 방식
1. 주석을 모두 제거한 페이지 사본에서 도면 원본 텍스트(문자 단위 좌표)를 읽습니다.
   - `GLV-0022` 처럼 한 줄인 Tag와, `SPV` / `3221B` 처럼 네모 박스 안에서 Prefix와 번호가 따로 있는 Tag를 모두 인식합니다. (회전된 텍스트 포함)
   - 결과 Tag는 `PREFIX-번호` 형식으로 통일됩니다 (`SPV 3221B` → `SPV-3221B`).
2. 각 Tag의 글자들이 어떤 주석에 가려졌는지 계산합니다.
   - 가리는 주석: 채우기색이 있는 Square/Circle/Polygon, FreeText(텍스트 박스), Redact, Stamp
   - 글자의 60% 이상이 가려지고 새 텍스트가 없으면 `DELETED`, 취소선(StrikeOut)도 `DELETED`
   - 가려진 부분 위에 텍스트 박스가 있으면 가려진 글자를 새 텍스트로 바꿔 `CHANGED`
3. 채우기 없는 구름/사각 표시 등은 Tag를 지우지 않은 것으로 보고 `Annotation` 열에만 기록합니다.

## 한계
- 주석이 이미 도면에 **Flatten(병합)** 된 PDF는 주석과 원본을 구분할 수 없습니다.
- Tag가 텍스트가 아니라 선(벡터)이나 이미지로 되어 있는 스캔 도면은 OCR이 필요합니다.
