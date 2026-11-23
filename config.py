# -*- coding: utf-8 -*-
"""
config.py  ―  모든 스크립트가 함께 쓰는 '설정' 파일입니다.

초보자 안내:
  - 인증키(API 키)는 이 파일에 직접 쓰지 말고, 같은 폴더의 .env 파일에 적으세요.
    (.env.example 파일을 복사해서 이름을 .env 로 바꾼 뒤 키를 붙여넣으면 됩니다.)
  - 아래 숫자(대기 시간, 최대 페이지 수 등)는 필요에 따라 바꿔도 됩니다.
"""

import os
from pathlib import Path
from urllib.parse import unquote

# data.go.kr 접속 시 "CERTIFICATE_VERIFY_FAILED"(보안인증서) 오류가 나는 경우가 많습니다.
# truststore 를 쓰면 파이썬이 윈도우에 설치된 인증서를 그대로 사용해서 이 오류가 사라집니다.
try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

# ─────────────────────────────────────────────
# 1. 폴더 위치
# ─────────────────────────────────────────────
# 이 파일(config.py)이 있는 폴더를 기준으로 모든 경로를 만듭니다.
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"            # 결과물을 모아두는 최상위 폴더
LIST_DIR = DATA_DIR / "lists"           # 1단계: 목록(Excel/CSV)
PDF_DIR = DATA_DIR / "pdf"              # 2단계: 내려받은 원문 파일
TEXT_DIR = DATA_DIR / "text"            # 3단계: 추출한 텍스트(.txt)
DB_PATH = DATA_DIR / "rda_docs.db"      # 4단계: SQLite 데이터베이스

# 1단계에서 모든 API 결과를 합친 목록 파일 (2·4단계가 이 파일을 읽습니다)
ALL_LIST_CSV = LIST_DIR / "목록_전체.csv"
# 2단계 다운로드 기록, 3단계 텍스트 추출 기록
DOWNLOAD_LOG_CSV = DATA_DIR / "다운로드_기록.csv"
EXTRACT_LOG_CSV = DATA_DIR / "텍스트추출_기록.csv"


# ─────────────────────────────────────────────
# 2. .env 파일에서 인증키 읽기
# ─────────────────────────────────────────────
def _load_env_file():
    """python-dotenv 가 있으면 그걸 쓰고, 없으면 직접 .env 를 한 줄씩 읽습니다."""
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return
    try:
        from dotenv import load_dotenv
        load_dotenv(env_path)
        return
    except ImportError:
        pass
    # python-dotenv 가 설치 안 된 경우를 위한 간단한 대체 방법
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


_load_env_file()


def get_service_key():
    """
    공공데이터포털 인증키를 돌려줍니다.
    포털에는 'Encoding 키'와 'Decoding 키' 두 가지가 있는데,
    이 프로그램은 Decoding 키를 기준으로 합니다.
    혹시 Encoding 키(%2B 같은 글자가 섞인 키)를 넣었다면 자동으로 풀어줍니다.
    """
    key = os.environ.get("DATA_GO_KR_SERVICE_KEY", "").strip()
    if not key or key.startswith("여기에"):
        raise SystemExit(
            "\n[오류] 인증키가 없습니다.\n"
            "  1) 공공데이터포털(data.go.kr)에서 API 활용신청을 하고 인증키를 받으세요.\n"
            "  2) 이 폴더의 .env.example 을 복사해 .env 로 이름을 바꾸세요.\n"
            "  3) .env 안의 DATA_GO_KR_SERVICE_KEY= 뒤에 Decoding 키를 붙여넣으세요.\n"
            "  자세한 방법은 README.md 의 '1. 인증키 받기'를 보세요.\n"
        )
    if "%" in key:  # Encoding 키를 넣은 경우 → Decoding 키로 변환
        key = unquote(key)
    return key


# ─────────────────────────────────────────────
# 3. 수집할 API 목록
# ─────────────────────────────────────────────
# 실제 호출 주소는 2026-11 기준으로 확인한 값입니다.
#   (주소가 바뀌면 오류코드 12 "해당 오픈API 서비스가 없거나 폐기됨"이 나옵니다.)
#
# 각 항목 설명
#   name          : 결과 파일 이름과 '출처' 칸에 쓰일 이름
#   list_url      : 목록 조회 주소
#   detail_url    : 상세 조회 주소 (목록에 첨부파일이 없을 때 하나씩 조회)
#   key_param     : 인증키를 넣는 파라미터 이름 (API마다 대소문자가 다릅니다)
#   page_param    : 페이지 번호 파라미터 이름
#   extra_params  : 그 외 항상 함께 보내는 값
APIS = {
    "연구동향": {
        "name": "국립원예특작과학원_연구동향",
        "list_url": "https://apis.data.go.kr/1390804/Nihhs_RstInfo/selectRstList",
        "detail_url": "https://apis.data.go.kr/1390804/Nihhs_RstInfo/selectRstView",
        "key_param": "serviceKey",
        "page_param": "pageNo",
        # 목록에 첨부파일 칸이 없으면 자료마다 상세 조회를 한 번 더 합니다 (시간이 더 걸림)
        "use_detail": True,
        "extra_params": {
            "numOfRows": 100,   # 한 번에 받을 건수
            "searchType": 1,    # 검색구분(필수값). 1 = 제목
            "searchWord": "",   # 비워두면 전체. 예: "월간리포트"
        },
    },
    "기술정보": {
        "name": "국립원예특작과학원_원예특작기술정보",
        "list_url": "https://apis.data.go.kr/1390804/Nihhs_TechInfo/selectTechInfoList",
        "detail_url": "https://apis.data.go.kr/1390804/Nihhs_TechInfo/selectTechInfoView",
        "key_param": "ServiceKey",
        "page_param": "currentPage",
        # 이 API는 목록에 본문과 첨부파일(atchfileUrl1~3)이 이미 들어 있어 상세 조회가 필요 없습니다
        "use_detail": False,
        "extra_params": {
            # 분야: ALL 전체, FT 과수, VC 채소, FL 화훼, IC 특용작물, MNG 경영, ETC 기타
            "ftrSeCode": "ALL",
            "searchType": 1,    # 1 제목, 2 작성자, 3 내용
            "searchWord": "",
        },
    },
}


# ─────────────────────────────────────────────
# 4. 예의 바른 수집(서버에 부담 주지 않기) 설정
# ─────────────────────────────────────────────
API_DELAY_SEC = 1.0        # API 호출 사이 대기 시간(초)
DOWNLOAD_DELAY_SEC = 3.0   # 파일 다운로드 사이 대기 시간(초). 줄이지 않는 것을 권장
MAX_PAGES = 500            # 혹시 모를 무한 반복 방지용 최대 페이지 수
REQUEST_TIMEOUT = 60       # 응답을 기다리는 최대 시간(초)
RETRY_COUNT = 3            # 실패 시 재시도 횟수

# 서버에 '누가 요청하는지' 알려주는 이름. 개인정보는 넣지 마세요.
# 2단계에서 내려받을 파일 종류 (첨부파일명 확장자 기준). 사진·동영상 등은 텍스트가 없어서 제외
# 첨부파일명에 확장자가 없으면 일단 받습니다.
DOWNLOAD_EXTS = (".pdf", ".hwp", ".hwpx")

USER_AGENT = "Mozilla/5.0 (research data collection script; python-requests)"

# 공공누리 라이선스는 API 응답에 들어있지 않습니다.
# 원문 페이지에서 확인한 뒤 Excel에서 직접 고쳐 쓰도록 기본값을 둡니다.
DEFAULT_LICENSE = "확인필요(공공누리 유형 확인)"


# ─────────────────────────────────────────────
# 5. 텍스트 추출(3단계) 설정
# ─────────────────────────────────────────────
# 한 페이지에서 뽑힌 글자가 이 수보다 적으면 '스캔본(이미지) PDF'로 봅니다.
MIN_CHARS_PER_PAGE = 30
# True 로 바꾸면 스캔본 PDF에 OCR(문자인식)을 시도합니다.
# Tesseract 프로그램과 한국어 데이터(kor)가 설치되어 있어야 합니다. README 참고.
USE_OCR = False
OCR_LANG = "kor+eng"
# Windows에서 tesseract.exe 위치 (기본 설치 경로). 다르면 고치세요.
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# ─────────────────────────────────────────────
# 6. RAG용 문단 나누기(4단계) 설정
# ─────────────────────────────────────────────
CHUNK_SIZE = 1000      # 한 조각의 최대 글자 수
CHUNK_OVERLAP = 150    # 조각끼리 겹치는 글자 수 (문맥이 끊기지 않게)


def make_folders():
    """결과 폴더들이 없으면 만들어 둡니다."""
    for folder in (DATA_DIR, LIST_DIR, PDF_DIR, TEXT_DIR):
        folder.mkdir(parents=True, exist_ok=True)
