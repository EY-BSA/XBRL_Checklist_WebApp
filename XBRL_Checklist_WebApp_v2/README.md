# XBRL 체크리스트 웹 애플리케이션

IxD 편집기 **구조내려받기 xlsx 파일**을 업로드하면 자동으로 XBRL 체크리스트를 검토

## 입력 파일

| 구분 | 형식 | 설명 |
|------|------|------|
| 구조내려받기 | `.xlsx` | IxD 편집기에서 내보낸 택사노미 구조 파일 |


## 구현된 체크 항목 (Alteryx 아웃풋 시트 번호 기준 — 29개)

| 번호 | 항목 | Alteryx Node |
|------|------|--------------|
| **1. 특정요소 사용검토** | | |
| 1-1 | Gross 계정 사용 검토 | Node 562→554→184 |
| 1-2 | 초과적립액(과소적립액) 텍사노미 사용 검토 | Node 563 |
| 1-3 | 재고자산 세부내역 표 검토 | Node 365→397→399 |
| 1-4 | 유동/비유동 축 검토 | Node 419 |
| **2. 텍사노미 검토** | | |
| 2-1 | (만료) 대손충당금 멤버 사용 | — |
| 2-2 | (만료) 금융자산 손상차손 축 | — |
| 2-3 | 대출약정 텍사노미 검토 | Node 388 |
| 2-4 | 미착품 텍사노미 검토 | Node 170→171→192 |
| **2-5** | **배당금 텍사노미 검토** ← 추가 | Node 604 |
| 2-6 | 평균유효세율 검토 (분반기) | Node 392 |
| **3. 축-멤버 정합성 검토** | | |
| 3-1 | Axis & Domain & Member 정합성 검토 | Node 324 |
| 3-2 | 공시금액의 사용 적정성 검토 | Node 468→593 |
| **4. 전용요소 사용 검토** | | |
| 4-1 ~ 4-6 | 현금흐름/판관비/특수관계자 전용요소 | Node 446~583 |
| **5. 속성/데이터타입 검토** | | |
| 5-1 | Percent 소숫점 자리수 검토 | Node 377 |
| 5-2 | 보유하는 주식수 속성 검토 | Node 483 |
| 5-3 | 이연법인세부채(자산) 검토 | Node 375 |
| 5-4 | 기본/희석주당이익 속성 검토 | Node 385 |
| **5-5** | **기초/기말 영문명 검토** ← 추가 | Node 597 |
| **5-6** | **단위표시 검토** ← 추가 | Node 619→620→621 |
| **6. 확장 검토** | | |
| 6-1 | 축 확장 검토 | Node 368→369 |
| 6-2 | 멤버 합계열 확장 검토 | Node 174→175 |
| **6-3** | **Duration / Instant 속성 검토** ← 추가 | Node 599 |
| **7. 기타** | | |
| 7-1 | Client Negate 검토 | Node 587 |
| 7-2 | 현금흐름표 영업활동 현금흐름 검토 | Node 423→428→424→429 |

## 5.12 수정내역

### taxonomy_xlsx_parser.py

 연결/별도 감지 로직 개선 | 기존 `role_def[6]` 방식 → 표 코드 마지막 글자 기준 (0=연결, 5=별도). `DX`, `DI` 접두사 코드에서 발생하던 오작동 해결 |
| 헤더 행 중복 제거 | 시트 내 반복 헤더행(`Name`, `Prefix` 등 컬럼명)이 데이터로 읽히지 않도록 필터 추가 |
| TABLE명 추적 | 각 시트 파싱 시 TABLE 행의 `Label(KO)`를 `table_name_ko`로 추적하여 모든 행에 추가 |
| Element 분류 로직 수정 | Alteryx 기준으로 순서 변경: ① suffix 기준 → ② `lineitem` 포함 시 `"Lineitem"` override → ③ `FOOTNOTES` 최종 override. 기존 `lineitem` 반환값 `"item"` → `"Lineitem"` 변경 |
| 축-도메인 그룹핑 필드 추가 | `_add_axis_group_fields()` 함수 추가. `축_도메인`, `Axis_flag`, `Axis_Name`, `GroupID`, `KEY_axis` 컬럼 생성 (3-1, 3-2 체크 전용) |
| axis_domain_rows 추가 | `GroupID`가 있는 행(축·도메인·멤버)만 추출하여 `data.axis_domain_rows`로 별도 저장 |

### checklist_engine.py

| 항목 | 내용 |
|------|------|
| DART_Negate_Check.xlsx 로드 | 모듈 시작 시 1회 로드. `Taxonomy ID`, `DART_Negate` 컬럼 사용. 7-1 체크에 활용 |
| Axis_Domain_Check.xlsx 로드 | 모듈 시작 시 1회 로드. `axis_check_records` 리스트로 사전 변환하여 성능 개선. 3-1 체크에 활용 |
| CheckIssue 필드 추가 | `client_negate`, `dart_negate` (7-1 전용), `table_name_ko` (전체 공통) |
| _mk() 수정 | `table_name_ko` 전달 추가 |

### app.py

| 항목 | 내용 |
|------|------|
| _serialize() 수정 | `table_name_ko` 직렬화 추가 (전체 체크). `client_negate`, `dart_negate` 추가 (7-1만) |

### templates/index.html

| 항목 | 내용 |
|------|------|
| TABLE명 컬럼 추가 | 전체 체크 결과에 TABLE명 컬럼 추가 |
| Client_Negate / DART_Negate 컬럼 추가 | 7-1 체크 결과에만 추가 |
| 컬럼 너비 조정 | `table-layout: fixed` 적용 및 컬럼별 고정 너비 지정. 7-1은 추가 컬럼을 고려한 별도 너비 적용 |

---

## 실행 방법

```bash
pip install -r requirements.txt
python app.py
# → http://localhost:5000
```

## 프로젝트 구조

```
XBRL_Checklist_WebApp/
├── app.py                      # Flask 웹 서버
├── taxonomy_xlsx_parser.py     # 구조내려받기 xlsx 파서
├── checklist_engine.py         # 체크리스트 검증 로직 (29개 체크)
├── requirements.txt
├── templates/
│   └── index.html              # 웹 UI (xlsx 드래그&드롭)
├── data/
│   ├── DART_Negate_Check.xlsx  # 7-1 Client Negate 검토용 참조 데이터
│   └── Axis_Domain_Check.xlsx  # 3-1 축-멤버 정합성 검토용 참조 데이터
└── 수정내역.xlsx                # 코드 수정 내역 정리
```

