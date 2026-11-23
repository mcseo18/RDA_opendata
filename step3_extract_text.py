# -*- coding: utf-8 -*-
"""
3단계: 내려받은 PDF에서 텍스트를 뽑아 .txt 파일로 저장합니다.

  방법
    1) PyMuPDF(fitz)로 먼저 추출 (빠르고 정확함)
    2) 실패하면 pdfplumber로 다시 시도
    3) 글자가 거의 안 나오는 페이지는 '스캔본(이미지)'으로 보고 표시합니다.
       config.py 에서 USE_OCR = True 로 바꾸면 Tesseract OCR(한국어)로 글자를 읽습니다.

  저장 위치
    data/text/<출처>/<파일이름>.txt     ← 페이지마다 "=== p.3 ===" 표시가 들어갑니다
    data/텍스트추출_기록.csv             ← 파일별 페이지 수, 글자 수, OCR 필요 여부

  실행 방법
    python step3_extract_text.py            → 아직 추출 안 한 PDF만 처리
    python step3_extract_text.py --redo     → 전부 다시 처리

  한글 파일(HWPX, HWP 5.0)도 글자를 뽑습니다. (페이지 구분은 없음)
"""

import argparse
import time
from pathlib import Path

import pandas as pd
from tqdm import tqdm

import config

LOG_COLUMNS = ["파일경로", "텍스트경로", "페이지수", "글자수", "추출방법", "OCR필요페이지", "상태"]


# ─────────────────────────────────────────────
# 추출 방법 1: PyMuPDF
# ─────────────────────────────────────────────
def extract_pymupdf(pdf_path):
    """페이지별 텍스트 목록을 돌려줍니다."""
    try:
        import pymupdf as fitz     # 최신 버전 이름
    except ImportError:
        import fitz                # 예전 버전 이름
    pages = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            pages.append(page.get_text("text") or "")
    return pages


# ─────────────────────────────────────────────
# 추출 방법 2: pdfplumber (대체용)
# ─────────────────────────────────────────────
def extract_pdfplumber(pdf_path):
    import pdfplumber
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return pages


# ─────────────────────────────────────────────
# 추출 방법 3: OCR (스캔본 페이지용, 선택)
# ─────────────────────────────────────────────
def ocr_pages(pdf_path, page_numbers):
    """지정한 페이지들을 그림으로 바꾼 뒤 Tesseract로 글자를 읽습니다."""
    try:
        import pytesseract
        from PIL import Image
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
    except ImportError:
        print("  ! OCR에는 pytesseract, Pillow 설치가 필요합니다: pip install pytesseract Pillow")
        return {}

    if Path(config.TESSERACT_CMD).exists():
        pytesseract.pytesseract.tesseract_cmd = config.TESSERACT_CMD

    results = {}
    with fitz.open(pdf_path) as doc:
        for no in page_numbers:
            pix = doc[no].get_pixmap(dpi=300)          # 300dpi 정도가 한글 인식에 적당
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            try:
                results[no] = pytesseract.image_to_string(img, lang=config.OCR_LANG)
            except pytesseract.TesseractError as err:
                print(f"  ! OCR 오류 (kor 언어 데이터가 설치됐는지 확인): {err}")
                return results
    return results


# ─────────────────────────────────────────────
# PDF 1개 처리
# ─────────────────────────────────────────────
def process_pdf(pdf_path):
    rel_pdf = str(pdf_path.relative_to(config.BASE_DIR))
    out_path = config.TEXT_DIR / pdf_path.relative_to(config.PDF_DIR)
    out_path = out_path.with_suffix(".txt")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    record = {"파일경로": rel_pdf, "텍스트경로": str(out_path.relative_to(config.BASE_DIR))}

    # 1) PyMuPDF → 2) pdfplumber 순서로 시도
    pages, method = None, ""
    for name, func in (("PyMuPDF", extract_pymupdf), ("pdfplumber", extract_pdfplumber)):
        try:
            pages, method = func(pdf_path), name
            break
        except Exception as err:   # 깨진 PDF 등 어떤 오류든 다음 방법으로 넘어감
            print(f"  ! {name} 실패 ({pdf_path.name}): {err}")
    if pages is None:
        record.update({"상태": "실패(PDF 열기 불가)", "페이지수": 0, "글자수": 0})
        return record

    # 3) 글자가 적은 페이지 = 스캔본으로 추정
    low = [i for i, t in enumerate(pages) if len(t.strip()) < config.MIN_CHARS_PER_PAGE]
    if low and config.USE_OCR:
        ocr_result = ocr_pages(pdf_path, low)
        for i, text in ocr_result.items():
            if len(text.strip()) > len(pages[i].strip()):
                pages[i] = text
        if ocr_result:
            method += "+OCR"
            low = [i for i in low if i not in ocr_result]

    # 페이지 번호를 붙여서 저장 (나중에 RAG에서 '몇 쪽'인지 알 수 있게)
    text = "\n\n".join(f"=== p.{i + 1} ===\n{t.strip()}" for i, t in enumerate(pages))
    out_path.write_text(text, encoding="utf-8")

    chars = sum(len(t.strip()) for t in pages)
    record.update({
        "페이지수": len(pages),
        "글자수": chars,
        "추출방법": method,
        "OCR필요페이지": ",".join(str(i + 1) for i in low[:50]) + (" ..." if len(low) > 50 else ""),
        "상태": "스캔본(OCR필요)" if pages and len(low) == len(pages) else "성공",
    })
    return record


# ─────────────────────────────────────────────
# 한글 파일(HWPX / HWP) 텍스트 추출
# ─────────────────────────────────────────────
def extract_hwpx(path):
    """
    HWPX는 사실 압축(zip) 파일이고, 안에 Contents/section0.xml 같은 글 내용이 들어 있습니다.
    <hp:p> = 문단, <hp:t> = 글자 조각 입니다. 페이지 정보는 없어서 1페이지로 취급합니다.
    """
    import zipfile
    import re as _re
    texts = []
    with zipfile.ZipFile(path) as z:
        sections = sorted(n for n in z.namelist()
                          if _re.match(r"Contents/section\d+\.xml$", n))
        if not sections:
            raise ValueError("HWPX 본문(section)을 찾지 못했습니다")
        for name in sections:
            xml = z.read(name).decode("utf-8", errors="ignore")
            for para in _re.findall(r"<hp:p[ >].*?</hp:p>", xml, flags=_re.S):
                line = "".join(_re.findall(r"<hp:t(?: [^>]*)?>(.*?)</hp:t>", para, flags=_re.S))
                line = _re.sub(r"<[^>]+>", "", line)
                texts.append(html_unescape(line))
    return ["\n".join(texts)]


def extract_hwp(path):
    """
    예전 한글 파일(HWP 5.0)에서 글자만 뽑습니다. olefile 패키지가 필요합니다.
    BodyText/Section0, Section1 ... 안의 '문단 글자' 기록(태그 번호 67)을 읽습니다.
    """
    import struct
    import zlib
    import olefile
    with olefile.OleFileIO(str(path)) as ole:
        header = ole.openstream("FileHeader").read()
        compressed = bool(header[36] & 1)
        sections = sorted(
            (e for e in ole.listdir() if e[0] == "BodyText"),
            key=lambda e: int(e[1].replace("Section", "") or 0),
        )
        paragraphs = []
        for entry in sections:
            data = ole.openstream(entry).read()
            if compressed:
                data = zlib.decompress(data, -15)
            pos = 0
            while pos + 4 <= len(data):
                head = struct.unpack_from("<I", data, pos)[0]
                tag, size = head & 0x3FF, (head >> 20) & 0xFFF
                pos += 4
                if size == 0xFFF:
                    size = struct.unpack_from("<I", data, pos)[0]
                    pos += 4
                if tag == 67:   # HWPTAG_PARA_TEXT
                    paragraphs.append(_hwp_para_text(data[pos:pos + size]))
                pos += size
    return ["\n".join(paragraphs)]


def _hwp_para_text(raw):
    """HWP 문단 글자에서 표·그림 같은 '조절 문자'를 건너뛰고 글자만 남깁니다."""
    out, i = [], 0
    n = len(raw) // 2
    while i < n:
        code = int.from_bytes(raw[i * 2:i * 2 + 2], "little")
        if code < 32:
            if code in (10, 13):
                out.append("\n")
            elif code == 9:
                out.append("\t")
            # 길이 8칸짜리 조절 문자들
            if code not in (0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31):
                i += 8
                continue
            i += 1
            continue
        out.append(chr(code))
        i += 1
    return "".join(out).strip()


def html_unescape(text):
    import html
    return html.unescape(text)


def process_hwp(path):
    """한글 파일 1개 처리 (PDF와 같은 모양의 기록을 돌려줌)."""
    rel = str(path.relative_to(config.BASE_DIR))
    out_path = (config.TEXT_DIR / path.relative_to(config.PDF_DIR)).with_suffix(".txt")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    record = {"파일경로": rel, "텍스트경로": str(out_path.relative_to(config.BASE_DIR))}
    is_zip = open(path, "rb").read(2) == b"PK"
    try:
        pages = extract_hwpx(path) if is_zip else extract_hwp(path)
        method = "HWPX" if is_zip else "HWP"
    except ImportError:
        record.update({"상태": "미지원(olefile 설치 필요: pip install olefile)", "텍스트경로": ""})
        return record
    except Exception as err:
        record.update({"상태": f"실패({str(err)[:60]})", "텍스트경로": ""})
        return record
    text = pages[0].strip()
    out_path.write_text(f"=== p.1 ===\n{text}", encoding="utf-8")
    record.update({"페이지수": 1, "글자수": len(text), "추출방법": method,
                   "상태": "성공" if text else "빈문서"})
    return record


def is_hwp(path):
    """HWP(구버전, OLE 형식) 또는 HWPX(zip 형식) 파일인지 확인합니다."""
    suffix = path.suffix.lower()
    try:
        with open(path, "rb") as f:
            head = f.read(4)
    except OSError:
        return False
    if suffix == ".hwpx" or (suffix != ".pdf" and head[:2] == b"PK"):
        return head[:2] == b"PK"
    return suffix == ".hwp" and head == b"\xd0\xcf\x11\xe0"


def is_pdf(path):
    """확장자가 달라도 내용이 PDF면 PDF로 봅니다."""
    try:
        with open(path, "rb") as f:
            return f.read(5) == b"%PDF-"
    except OSError:
        return False


def main():
    parser = argparse.ArgumentParser(description="PDF 텍스트 추출")
    parser.add_argument("--redo", action="store_true", help="이미 추출한 것도 다시 처리")
    args = parser.parse_args()
    config.make_folders()

    files = [p for p in config.PDF_DIR.rglob("*") if p.is_file() and not p.name.endswith(".part")]
    if not files:
        raise SystemExit("data/pdf 폴더에 파일이 없습니다. 먼저 step2_download_pdf.py 를 실행하세요.")

    old = pd.DataFrame(columns=LOG_COLUMNS)
    if config.EXTRACT_LOG_CSV.exists():
        old = pd.read_csv(config.EXTRACT_LOG_CSV, dtype=str, keep_default_na=False)
    done = set() if args.redo else set(old["파일경로"])

    records = []
    start = time.time()
    for path in tqdm(files, desc="텍스트 추출"):
        rel = str(path.relative_to(config.BASE_DIR))
        if rel in done:
            continue
        if is_pdf(path):
            records.append(process_pdf(path))
        elif is_hwp(path):
            records.append(process_hwp(path))
        else:
            records.append({"파일경로": rel, "상태": f"미지원({path.suffix or '확장자없음'})"})

    new = pd.DataFrame(records, columns=LOG_COLUMNS)
    merged = pd.concat([old[~old["파일경로"].isin(new["파일경로"])], new], ignore_index=True)
    merged.to_csv(config.EXTRACT_LOG_CSV, index=False, encoding="utf-8-sig")

    scanned = (new["상태"] == "스캔본(OCR필요)").sum() if len(new) else 0
    print(f"\n완료: {len(new)}개 처리, {time.time() - start:.0f}초 → 기록: {config.EXTRACT_LOG_CSV}")
    if scanned and not config.USE_OCR:
        print(f"스캔본으로 보이는 PDF {scanned}개가 있습니다. OCR이 필요하면 README의 'OCR' 부분을 보세요.")
    print("다음 단계: python step4_build_db.py")


if __name__ == "__main__":
    main()
