# 농촌진흥청 자료 수집 → RAG 준비 도구

공공데이터포털(data.go.kr) API로 국립원예특작과학원 자료 **목록을 모으고**, 원문 **PDF를 내려받고**,
**텍스트를 뽑아서**, 나중에 RAG(검색 기반 AI 답변)에 쓸 수 있도록 **SQLite 데이터베이스 하나로 정리**합니다.

```
1단계 목록 수집 ──▶ 2단계 PDF 다운로드 ──▶ 3단계 텍스트 추출 ──▶ 4단계 DB 통합
 (Excel/CSV)          (data/pdf)             (data/text)            (data/rda_docs.db)
```

---

## 0. 폴더 구성

| 파일 | 하는 일 |
|---|---|
| `config.py` | 모든 설정(폴더 위치, API 주소, 대기 시간, OCR 사용 여부 등) |
| `.env.example` | 인증키를 적는 파일의 견본 (복사해서 `.env` 로 이름 변경) |
| `step1_collect_list.py` | 1단계: API로 목록 수집 → Excel/CSV |
| `step2_download_pdf.py` | 2단계: 목록의 URL에서 원문 파일 내려받기 (파일마다 3초 쉬기) |
| `step3_extract_text.py` | 3단계: PDF에서 텍스트 추출 (스캔본은 OCR 선택) |
| `step4_build_db.py` | 4단계: SQLite로 합치기, 발간등록번호로 중복 제거, RAG용 조각 만들기 |
| `requirements.txt` | 필요한 파이썬 패키지 목록 |

실행하면 `data` 폴더가 생기고 결과물이 모두 그 안에 쌓입니다.

---

## 1. 인증키 받기 (가장 먼저, 한 번만)

지금은 인증키가 없으므로 이것부터 해야 합니다. 승인까지 보통 바로~1시간 정도 걸립니다.

1. [공공데이터포털](https://www.data.go.kr) 회원가입 후 로그인
2. 아래 두 API 페이지에 들어가서 각각 **[활용신청]** 버튼을 누릅니다.
   - 연구동향: <https://www.data.go.kr/data/15135039/openapi.do>
     (농촌진흥청 국립원예특작과학원_원예특작 연구동향 조회 서비스)
   - 기술정보: <https://www.data.go.kr/data/15040630/openapi.do>
     (농촌진흥청 국립원예특작과학원_원예특용작물 기술정보 조회 서비스)
3. 활용목적은 "연구(논문 등)" 또는 "참고자료" 선택, 내용에 예) *"원예·특용작물 연구자료 수집 및 내부 연구용 검색 시스템 구축"* 정도로 적습니다.
4. **마이페이지 → 데이터활용 → Open API → 활용신청 현황**에서 승인된 것을 확인하고,
   **일반 인증키(Decoding)** 를 복사합니다.
   - 인증키는 계정에 하나라서 두 API 모두 같은 키를 씁니다.
   - Encoding 키를 넣어도 프로그램이 알아서 바꿔주지만, Decoding 키를 권장합니다.
5. 이 폴더의 `.env.example` 을 복사해서 이름을 `.env` 로 바꾸고, 키를 붙여넣습니다.

```powershell
copy .env.example .env
notepad .env
```

```
DATA_GO_KR_SERVICE_KEY=복사한_Decoding_키
```

> ⚠️ `.env` 파일은 비밀번호와 같습니다. 다른 사람에게 보내거나 GitHub 등에 올리지 마세요.
> 승인 직후 "등록되지 않은 서비스키(오류 30)"가 나오면 1~2시간 뒤 다시 해보세요.

---

## 2. 설치 (한 번만)

PowerShell에서 이 폴더로 이동한 뒤, 이 프로젝트 전용 가상환경을 만들어 설치하는 것을 권장합니다.

```powershell
cd D:\RDA_opendata
```

```powershell
python -m venv .venv
```

```powershell
.\.venv\Scripts\Activate.ps1
```

```powershell
pip install -r requirements.txt
```

- 다음부터는 작업할 때마다 `.\.venv\Scripts\Activate.ps1` 만 실행하면 됩니다 (앞에 `(.venv)` 표시가 붙음).
- "스크립트를 실행할 수 없습니다" 오류가 나면 한 번만 실행:
  `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

---

## 3. 실행 순서

### 1단계: 목록 수집

처음에는 **테스트 모드**로 응답 모양부터 확인하세요 (1페이지만 받아서 화면에 보여줌).

```powershell
python step1_collect_list.py --test
```

문제가 없으면 전체 수집:

```powershell
python step1_collect_list.py
```

- 결과: `data\lists\목록_연구동향.xlsx`, `목록_기술정보.xlsx`, `목록_전체.csv`
- 칸: **제목, 발간년도, 발간등록번호, URL, 라이선스** + 첨부파일명, 원문ID, 작성일, 발행기관, 분류, 원문페이지, 본문
  - 첨부파일이 여러 개인 자료는 파일마다 한 줄씩 나옵니다.
  - **기술정보 API**는 본문(`contentsKor`)이 함께 오므로, 첨부파일이 없어도 본문이 RAG 자료로 쓰입니다.
  - **연구동향 API**는 목록에 첨부파일이 없으면 자료마다 상세 조회를 한 번 더 해서 다운로드 URL을 찾습니다 (그래서 조금 느립니다).
  - **발간등록번호**는 API가 따로 주지 않아서, 제목·본문·파일명에서 `11-1390000-000000-01` 모양을 찾아 채웁니다. 4단계에서 PDF 본문에서도 한 번 더 찾습니다.
  - **라이선스**는 API에 없어서 `확인필요`로 채워집니다. 아래 7번 참고.
- 기술정보 분야만 받고 싶으면 `config.py` 의 `ftrSeCode` 를 바꾸세요 (FT 과수, VC 채소, FL 화훼, IC 특용작물).
- 한 API만: `python step1_collect_list.py --only 기술정보`

### 2단계: 원문 파일 내려받기

시험 삼아 10개만 먼저 받아보세요.

```powershell
python step2_download_pdf.py --limit 10
```

```powershell
python step2_download_pdf.py
```

- 파일 하나마다 **3초씩 쉬면서** 받습니다 (`config.py` 의 `DOWNLOAD_DELAY_SEC`). 서버에 무리가 가지 않도록 줄이지 않는 것을 권장합니다. 1,000개면 1시간 가까이 걸립니다.
- 중간에 `Ctrl + C` 로 멈춰도 괜찮습니다. 다시 실행하면 **이미 받은 것은 건너뛰고** 이어서 받습니다.
- 실패한 것만 다시: `python step2_download_pdf.py --retry-failed`
- 기록: `data\다운로드_기록.csv` (파일형식 칸에 pdf / hwp 등이 표시됨)

### 3단계: 텍스트 추출

```powershell
python step3_extract_text.py
```

- PyMuPDF로 먼저 뽑고, 안 되면 pdfplumber로 다시 시도합니다.
- 결과: `data\text\...\파일이름.txt` (페이지마다 `=== p.3 ===` 표시 → 나중에 "몇 쪽" 출처 표시에 사용)
- 기록: `data\텍스트추출_기록.csv` → **상태가 `스캔본(OCR필요)`** 인 파일은 글자가 아니라 그림으로 된 PDF입니다. 아래 5번 OCR 참고.

### 4단계: 데이터베이스로 합치기

```powershell
python step4_build_db.py
```

- 결과: `data\rda_docs.db` (SQLite), `data\최종목록.xlsx`
- **중복 제거**: 발간등록번호가 같으면 같은 자료로 보고 하나만 남깁니다. 번호가 없으면 URL로, 그것도 없으면 출처+원문ID+제목으로 구분합니다. 내용이 글자 하나 안 틀리고 같은 파일도 하나만 남깁니다.
- DB 안의 표
  - `documents`: 자료 1건 = 1행 (메타데이터 + 전문)
  - `chunks`: RAG용으로 약 1,000자씩 자른 조각 (어느 페이지인지 포함)
  - `chunks_fts`: 키워드 검색 색인
- DB 내용은 무료 프로그램 [DB Browser for SQLite](https://sqlitebrowser.org/) 로 열어볼 수 있습니다.
- 여러 번 실행해도 됩니다. DB에서 직접 고친 **라이선스** 값은 지워지지 않습니다.

---

## 4. 처음부터 끝까지 한 번에 (익숙해진 뒤)

```powershell
python step1_collect_list.py; python step2_download_pdf.py; python step3_extract_text.py; python step4_build_db.py
```

---

## 5. 스캔본 PDF 글자 읽기 (OCR, 선택)

오래된 보고서는 종이를 스캔한 PDF라 글자가 안 뽑힙니다. 이때 Tesseract OCR을 씁니다.

1. Windows용 Tesseract 설치: <https://github.com/UB-Mannheim/tesseract/wiki>
   - 설치 중 **Additional language data** 에서 **Korean** 을 꼭 체크하세요 (`kor` 데이터).
   - 기본 설치 위치: `C:\Program Files\Tesseract-OCR\` (다르면 `config.py` 의 `TESSERACT_CMD` 수정)
2. 파이썬 연결 패키지 설치:
   ```powershell
   pip install pytesseract Pillow
   ```
3. `config.py` 에서 `USE_OCR = True` 로 바꾸고, 다시 추출:
   ```powershell
   python step3_extract_text.py --redo
   ```

참고: OCR은 페이지당 수 초씩 걸리고, 표나 세로쓰기는 정확도가 떨어집니다. 중요한 자료는 결과를 눈으로 확인하세요.

## 6. HWP 파일

일부 첨부파일은 PDF가 아니라 한글(HWP/HWPX) 파일입니다. 3단계는 이를 `미지원`으로 기록만 합니다.
양이 적으면 한글 프로그램에서 **PDF로 저장**해 `data\pdf\` 아래 같은 폴더에 넣고 3단계를 다시 실행하는 것이 가장 간단합니다.

---

## 7. 라이선스(공공누리) 주의사항

공공기관 저작물에는 보통 **공공누리** 표시가 붙어 있고, 유형마다 허용 범위가 다릅니다.

| 유형 | 조건 | 의미 |
|---|---|---|
| 제1유형 | 출처표시 | 상업적 이용, 변경 모두 가능 |
| 제2유형 | 출처표시 + 상업적 이용금지 | 비영리 목적만 |
| 제3유형 | 출처표시 + 변경금지 | 내용을 고치거나 2차 저작물 작성 불가 |
| 제4유형 | 출처표시 + 상업적 이용금지 + 변경금지 | 가장 제한적 |

- API 응답에는 라이선스 정보가 없어서 `확인필요`로 들어갑니다. 각 자료의 **원문페이지**(또는 PDF 판권면)에서 공공누리 마크를 확인해 `최종목록.xlsx`나 DB의 `라이선스` 칸을 채워 두세요.
- 개인·연구실 내부 연구용 검색에 쓰는 것은 대체로 문제가 없지만, **RAG 서비스를 외부에 공개하거나 유료로 제공**할 계획이면 제2·4유형(상업적 금지), 제3·4유형(변경금지 → 요약·재가공 답변 주의)에 해당하는 자료를 따로 확인하세요.
- AI 답변에는 **출처(제목, 발행기관, 쪽수)** 를 함께 보여주도록 만들면 출처표시 조건을 지키기 쉽습니다. DB의 `chunks.페이지` 칸이 이 용도입니다.
- 공공데이터포털 API 자체의 이용허락범위는 각 API 상세페이지에 적혀 있습니다.

> 이 안내는 일반적인 참고사항이며 법률 자문이 아닙니다. 공개 서비스로 만들 때는 기관 담당자나 공공누리 안내(<https://www.kogl.or.kr>)를 확인하세요.

---

## 8. 더 많은 자료를 모으는 방법

### 농사로 오픈API
농촌진흥청 [농사로](https://www.nongsaro.go.kr)는 영농기술, 농업기술길잡이, 주간농사정보, 작목정보 등을 별도 오픈API로 제공합니다.
- 농사로 회원가입 → **오픈API** 메뉴에서 사용 신청 → **농사로 전용 인증키(apiKey)** 를 받습니다. (공공데이터포털 키와 다릅니다)
- 키를 받으면 `.env` 에 `NONGSARO_API_KEY=...` 로 따로 저장하고, `config.py` 의 `APIS` 에 같은 모양으로 항목을 추가하면 1단계 스크립트로 수집할 수 있습니다. (농사로 API는 인증키 파라미터 이름이 `apiKey` 이므로 `key_param` 을 그에 맞게 지정)

### 정보공개청구
API에 없는 보고서(예: 연구사업 최종보고서, 오래된 발간자료 원문)는 [정보공개포털](https://www.open.go.kr)에서 **농촌진흥청(또는 국립원예특작과학원)** 을 대상으로 청구할 수 있습니다.
- 청구 예시: *"20XX~20XX년 국립원예특작과학원 발간 연구보고서 목록(제목, 발간등록번호, 발간년도) 및 전자파일 사본"*
- 원칙적으로 10일 이내 결정(한 번 10일 연장 가능)되며, 전자파일로 받으면 `data\pdf\정보공개청구\` 폴더에 넣고 3·4단계를 실행하면 됩니다. (이 경우 목록 CSV에 행을 직접 추가해야 DB에 메타데이터가 붙습니다)

또한 농촌진흥청 농업과학도서관, [NTIS](https://www.ntis.go.kr), [ScienceON](https://scienceon.kisti.re.kr) 에서도 농진청 보고서 원문을 찾을 수 있습니다. 웹사이트를 자동으로 긁어오는 것(크롤링)은 각 사이트 이용약관과 robots.txt를 먼저 확인하세요.

---

## 9. 자주 나오는 오류

| 화면에 나온 말 | 원인과 해결 |
|---|---|
| `인증키가 없습니다` | `.env` 파일이 없거나 키가 비어 있음 → 1번 다시 확인 |
| `[API 오류 30] 등록되지 않은 서비스키` | 키 오타, 또는 승인 직후 → 1~2시간 뒤 재시도 |
| `[API 오류 20]` | 해당 API 활용신청이 안 됨 → 두 API 각각 신청했는지 확인 |
| `[API 오류 22]` | 하루 호출 한도 초과 → 내일 다시 (이미 받은 것은 남아 있음) |
| `[API 오류 12]` | API 주소가 바뀜 → 상세페이지에서 주소 확인 후 `config.py` 수정 |
| `CERTIFICATE_VERIFY_FAILED` | 보안인증서 오류 → `pip install truststore` (requirements에 포함됨) |
| `ModuleNotFoundError` | 패키지 미설치, 또는 가상환경을 안 켬 → `.\.venv\Scripts\Activate.ps1` 후 `pip install -r requirements.txt` |
| `PermissionError ... xlsx` | 결과 엑셀 파일이 열려 있음 → 닫고 다시 실행 |
| 1단계 결과에 제목이 비어 있음 | API 칸 이름이 예상과 다름 → `--test` 결과의 칸 이름을 `step1_collect_list.py` 위쪽 `TITLE_KEYS` 등에 추가 |

---

## 10. 다음 단계 (RAG)

`data\rda_docs.db` 의 `chunks` 표가 RAG의 재료입니다. 각 조각(`내용`)을 임베딩(벡터)으로 바꿔 벡터DB(Chroma, FAISS 등)에 넣고,
질문이 오면 비슷한 조각을 찾아 `documents` 표의 제목·발행기관·페이지와 함께 AI에게 넘기면 됩니다.
