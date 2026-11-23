# -*- coding: utf-8 -*-
"""
2단계: 1단계 목록(data/lists/목록_전체.csv)의 URL에서 원문 파일(PDF 등)을 내려받습니다.

  ★ 서버에 부담을 주지 않도록 파일 하나 받을 때마다 쉬었다가(time.sleep) 받습니다.
    (config.py 의 DOWNLOAD_DELAY_SEC, 기본 3초)
  ★ 이미 받은 파일은 건너뛰므로, 중간에 멈춰도 다시 실행하면 이어서 받습니다.

  저장 위치
    data/pdf/<출처>/<원문ID>_<파일명>
    data/다운로드_기록.csv   ← URL별 저장 경로와 성공/실패 기록 (4단계가 읽음)

  실행 방법
    python step2_download_pdf.py              → 전부 받기
    python step2_download_pdf.py --limit 20   → 시험 삼아 20개만 받기
"""

import argparse
import re
import time
from pathlib import Path
from urllib.parse import unquote, urlparse

import pandas as pd
import requests
from tqdm import tqdm

import config

LOG_COLUMNS = ["URL", "파일경로", "상태", "크기(KB)", "파일형식", "시각"]


def safe_name(text, max_len=80):
    """윈도우 파일 이름에 쓸 수 없는 글자(\\ / : * ? " < > |)를 _ 로 바꿉니다."""
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(text)).strip(" ._")
    return text[:max_len] or "noname"


def guess_file_name(resp, url, fallback):
    """서버가 알려주는 파일 이름(Content-Disposition)을 우선 쓰고, 없으면 URL에서 찾습니다."""
    cd = resp.headers.get("Content-Disposition", "")
    match = re.search(r"filename\*=UTF-8''([^;]+)", cd, re.I) or \
        re.search(r'filename="?([^";]+)"?', cd, re.I)
    if match:
        raw = match.group(1)
        name = unquote(raw)
        # 한글 파일명이 깨져서 오는 경우(latin-1 로 잘못 해석) 복구 시도
        for enc in ("utf-8", "cp949"):
            try:
                name = raw.encode("latin-1").decode(enc)
                break
            except (UnicodeEncodeError, UnicodeDecodeError):
                continue
        return unquote(name)
    tail = Path(unquote(urlparse(url).path)).name
    return tail if "." in tail else fallback


def detect_type(first_bytes, content_type):
    """파일 앞부분을 보고 실제 형식을 판단합니다."""
    if first_bytes.startswith(b"%PDF"):
        return "pdf"
    if first_bytes.startswith(b"PK"):
        return "zip/hwpx/docx"
    if first_bytes.startswith(b"\xd0\xcf\x11\xe0"):
        return "hwp/doc(구버전)"
    if b"<html" in first_bytes[:500].lower() or "html" in content_type:
        return "html(오류페이지일 수 있음)"
    return "기타"


def load_log():
    if config.DOWNLOAD_LOG_CSV.exists():
        return pd.read_csv(config.DOWNLOAD_LOG_CSV, dtype=str, keep_default_na=False)
    return pd.DataFrame(columns=LOG_COLUMNS)


def download_one(session, row):
    """파일 1개를 내려받고 기록(딕셔너리)을 돌려줍니다."""
    url = row["URL"]
    folder = config.PDF_DIR / safe_name(row["출처"])
    folder.mkdir(parents=True, exist_ok=True)
    prefix = safe_name(row.get("원문ID") or "x", 20)
    record = {"URL": url, "시각": time.strftime("%Y-%m-%d %H:%M:%S")}

    for attempt in range(1, config.RETRY_COUNT + 1):
        try:
            with session.get(url, stream=True, timeout=config.REQUEST_TIMEOUT) as resp:
                resp.raise_for_status()
                fallback = safe_name(row.get("첨부파일명") or row["제목"]) + ".pdf"
                name = safe_name(row.get("첨부파일명") or guess_file_name(resp, url, fallback), 100)
                path = folder / f"{prefix}_{name}"
                tmp = path.with_suffix(path.suffix + ".part")   # 받는 중 임시 파일
                first = b""
                with open(tmp, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=64 * 1024):
                        if not first:
                            first = chunk[:1024]
                        f.write(chunk)
                tmp.replace(path)   # 다 받으면 진짜 이름으로 바꿈
                ftype = detect_type(first, resp.headers.get("Content-Type", ""))
                record.update({
                    "파일경로": str(path.relative_to(config.BASE_DIR)),
                    "상태": "성공" if ftype != "html(오류페이지일 수 있음)" else "확인필요",
                    "크기(KB)": round(path.stat().st_size / 1024, 1),
                    "파일형식": ftype,
                })
                return record
        except (requests.RequestException, OSError) as err:
            record.update({"상태": f"실패: {str(err)[:100]}", "파일경로": ""})
            time.sleep(config.DOWNLOAD_DELAY_SEC * attempt)
    return record


def main():
    parser = argparse.ArgumentParser(description="원문 파일 일괄 다운로드")
    parser.add_argument("--limit", type=int, default=0, help="이번에 받을 최대 개수 (0 = 전부)")
    parser.add_argument("--retry-failed", action="store_true", help="예전에 실패한 것도 다시 시도")
    args = parser.parse_args()

    if not config.ALL_LIST_CSV.exists():
        raise SystemExit("목록 파일이 없습니다. 먼저 python step1_collect_list.py 를 실행하세요.")
    config.make_folders()

    df = pd.read_csv(config.ALL_LIST_CSV, dtype=str, keep_default_na=False)
    df = df[df["URL"].str.startswith("http")].drop_duplicates(subset="URL")
    # 사진·동영상 등 텍스트로 못 바꾸는 파일은 건너뜀 (config.DOWNLOAD_EXTS)
    names = df["첨부파일명"].str.lower()
    has_ext = names.str.contains(r"\.[a-z0-9]{2,5}$", regex=True)
    skip = has_ext & ~names.str.endswith(config.DOWNLOAD_EXTS)
    if skip.any():
        print(f"텍스트 추출이 안 되는 파일 {skip.sum()}개는 건너뜀: {sorted(set(names[skip].str.extract(r'(\.[a-z0-9]+)$')[0]))}")
    df = df[~skip]

    log = load_log()
    done = set(log.loc[log["상태"] == "성공", "URL"])
    if not args.retry_failed:
        done |= set(log.loc[log["상태"].str.startswith("실패"), "URL"])
    todo = df[~df["URL"].isin(done)]
    if args.limit:
        todo = todo.head(args.limit)

    print(f"받을 파일 {len(todo)}개 (이미 처리한 것 {len(done)}개는 건너뜀)")
    print(f"파일 사이 대기 {config.DOWNLOAD_DELAY_SEC}초 → 예상 최소 {len(todo) * config.DOWNLOAD_DELAY_SEC / 60:.0f}분")

    session = requests.Session()
    session.headers["User-Agent"] = config.USER_AGENT
    new_records = []
    try:
        for _, row in tqdm(todo.iterrows(), total=len(todo), desc="다운로드"):
            new_records.append(download_one(session, row))
            time.sleep(config.DOWNLOAD_DELAY_SEC)   # ★ 서버 배려용 대기
            if len(new_records) % 20 == 0:          # 20개마다 중간 저장
                save_log(log, new_records)
    except KeyboardInterrupt:
        print("\n사용자가 멈췄습니다. 지금까지 받은 기록을 저장합니다.")
    save_log(log, new_records)

    ok = sum(r["상태"] == "성공" for r in new_records)
    print(f"\n완료: 성공 {ok}개 / 이번 시도 {len(new_records)}개 → 기록: {config.DOWNLOAD_LOG_CSV}")
    print("다음 단계: python step3_extract_text.py")


def save_log(old_log, new_records):
    if not new_records:
        return
    new = pd.DataFrame(new_records, columns=LOG_COLUMNS)
    merged = pd.concat([old_log[~old_log["URL"].isin(new["URL"])], new], ignore_index=True)
    merged.to_csv(config.DOWNLOAD_LOG_CSV, index=False, encoding="utf-8-sig")


if __name__ == "__main__":
    main()
