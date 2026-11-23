# -*- coding: utf-8 -*-
"""
4단계: 목록 + 다운로드 기록 + 추출 텍스트를 SQLite 데이터베이스 하나로 합칩니다.

  중복 제거 기준 (위에서부터 먼저 있는 값을 씀)
    1) 발간등록번호  → 같은 번호는 같은 자료로 봄
    2) URL           → 번호가 없으면 다운로드 주소로 구분
    3) 출처+원문ID+제목
  그리고 내용(텍스트)이 완전히 같은 파일이 두 번 들어온 경우도 하나만 남깁니다.

  만들어지는 것 (data 폴더)
    rda_docs.db         ← SQLite 데이터베이스
       - documents 표 : 자료 1건 = 1행 (메타데이터 + 전문)
       - chunks 표    : RAG용으로 전문을 약 1,000자씩 자른 조각 (페이지 번호 포함)
       - chunks_fts   : 조각 키워드 검색용 색인 (지원되는 경우)
    최종목록.xlsx       ← documents 표를 엑셀로 (전문 제외)

  실행 방법
    python step4_build_db.py
    여러 번 실행해도 됩니다. 이미 있는 자료는 새 정보로 갱신됩니다.
    (DB에서 직접 고친 '라이선스' 값은 기본값으로 덮어쓰지 않습니다.)

  DB 내용 보기: 'DB Browser for SQLite' 프로그램(무료)으로 rda_docs.db 를 열면 됩니다.
"""

import hashlib
import re
import sqlite3
import time
from pathlib import Path

import pandas as pd

import config

SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
    doc_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    중복키         TEXT UNIQUE NOT NULL,
    출처           TEXT,
    제목           TEXT,
    발간년도       TEXT,
    발간등록번호   TEXT,
    URL            TEXT,
    라이선스       TEXT,
    첨부파일명     TEXT,
    원문ID         TEXT,
    작성일         TEXT,
    발행기관       TEXT,
    분류           TEXT,
    원문페이지     TEXT,
    파일경로       TEXT,
    텍스트경로     TEXT,
    추출방법       TEXT,
    글자수         INTEGER,
    내용해시       TEXT,
    전문           TEXT,
    수정일시       TEXT
);
CREATE TABLE IF NOT EXISTS chunks (
    chunk_id  INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_id    INTEGER REFERENCES documents(doc_id) ON DELETE CASCADE,
    순번      INTEGER,
    페이지    INTEGER,
    내용      TEXT
);
CREATE INDEX IF NOT EXISTS idx_chunks_doc ON chunks(doc_id);
CREATE INDEX IF NOT EXISTS idx_docs_pubno ON documents(발간등록번호);
"""

PUB_NO_PATTERN = re.compile(r"\b\d{2}-\d{7}-\d{6}-\d{2}\b")
PAGE_MARK = re.compile(r"=== p\.(\d+) ===")


def read_csv(path):
    """CSV가 있으면 읽고, 없으면 빈 표를 돌려줍니다."""
    if Path(path).exists():
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    return pd.DataFrame()


def make_key(row):
    """중복 판단용 키를 만듭니다."""
    if row.get("발간등록번호"):
        return "PUB:" + row["발간등록번호"]
    if row.get("URL"):
        return "URL:" + row["URL"]
    return "ID:" + "|".join([row.get("출처", ""), row.get("원문ID", ""), row.get("제목", "")])


def split_chunks(text):
    """
    전문을 RAG용 조각으로 나눕니다.
    '=== p.3 ===' 표시로 페이지를 알아내고, 문단 단위로 CHUNK_SIZE 글자까지 모읍니다.
    돌려주는 값: [(페이지번호, 조각내용), ...]
    """
    # 1) 페이지별로 나누기 (표시가 없으면 전체를 1페이지로)
    parts = PAGE_MARK.split(text)
    if len(parts) == 1:
        pages = [(1, text)]
    else:
        pages = [(int(parts[i]), parts[i + 1]) for i in range(1, len(parts) - 1, 2)]

    # 2) 문단 단위로 쪼갠 뒤 정해진 길이까지 이어 붙이기
    chunks, buf, buf_page = [], "", None
    for page_no, page_text in pages:
        for para in re.split(r"\n\s*\n", page_text):
            para = para.strip()
            if not para:
                continue
            while len(para) > config.CHUNK_SIZE:          # 너무 긴 문단은 강제로 자름
                if buf:
                    chunks.append((buf_page, buf))
                    buf = ""
                chunks.append((page_no, para[:config.CHUNK_SIZE]))
                para = para[config.CHUNK_SIZE - config.CHUNK_OVERLAP:]
            if buf and len(buf) + len(para) + 1 > config.CHUNK_SIZE:
                chunks.append((buf_page, buf))
                tail = buf[-config.CHUNK_OVERLAP:]        # 앞 조각 끝부분을 겹쳐서 문맥 유지
                buf, buf_page = tail + "\n" + para, page_no
            else:
                buf_page = buf_page if buf else page_no
                buf = (buf + "\n" + para) if buf else para
    if buf.strip():
        chunks.append((buf_page, buf))
    return chunks


def build_table():
    """세 가지 CSV를 합쳐서 '자료 1건 = 1행' 표를 만듭니다."""
    lists = read_csv(config.ALL_LIST_CSV)
    if lists.empty:
        raise SystemExit("목록 파일이 없습니다. 먼저 python step1_collect_list.py 를 실행하세요.")
    downloads = read_csv(config.DOWNLOAD_LOG_CSV)
    extracts = read_csv(config.EXTRACT_LOG_CSV)

    # URL → 내려받은 파일경로
    if not downloads.empty:
        ok = downloads[downloads["상태"] == "성공"][["URL", "파일경로"]]
        lists = lists.merge(ok, on="URL", how="left")
    else:
        lists["파일경로"] = ""
    # 파일경로 → 텍스트경로, 추출방법
    if not extracts.empty:
        lists = lists.merge(extracts[["파일경로", "텍스트경로", "추출방법"]], on="파일경로", how="left")
    else:
        lists["텍스트경로"], lists["추출방법"] = "", ""
    return lists.fillna("")


def main():
    config.make_folders()
    table = build_table()
    print(f"목록 {len(table)}행을 DB로 합칩니다...")

    conn = sqlite3.connect(config.DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA)
    try:   # 키워드 검색 색인 (SQLite FTS5). 지원 안 되면 건너뜀
        conn.execute("CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(내용, content='chunks', content_rowid='chunk_id')")
        has_fts = True
    except sqlite3.OperationalError:
        has_fts = False

    now = time.strftime("%Y-%m-%d %H:%M:%S")
    seen_keys, seen_hashes = set(), {}
    stats = {"추가/갱신": 0, "키중복": 0, "내용중복": 0}

    for _, row in table.iterrows():
        row = row.to_dict()

        # 전문 정하기: PDF에서 뽑은 텍스트가 있으면 그것, 없으면 API 본문(contentsKor)
        full_text, method = "", row.get("추출방법", "")
        txt_path = config.BASE_DIR / row["텍스트경로"] if row.get("텍스트경로") else None
        if txt_path and txt_path.exists():
            full_text = txt_path.read_text(encoding="utf-8")
        if len(PAGE_MARK.sub("", full_text).strip()) < 50 and row.get("본문"):
            full_text, method = row["본문"], "API본문"

        # 목록에 발간등록번호가 없으면 전문 앞부분에서 찾아봄 (보통 판권면/표지에 있음)
        if not row.get("발간등록번호"):
            m = PUB_NO_PATTERN.search(full_text[:20000]) or PUB_NO_PATTERN.search(full_text[-5000:])
            row["발간등록번호"] = m.group(0) if m else ""

        key = make_key(row)
        if key in seen_keys:
            stats["키중복"] += 1
            continue
        seen_keys.add(key)

        body = PAGE_MARK.sub("", full_text).strip()
        digest = hashlib.sha1(re.sub(r"\s+", "", body).encode("utf-8")).hexdigest() if body else ""
        if digest and digest in seen_hashes:
            stats["내용중복"] += 1
            continue
        if digest:
            seen_hashes[digest] = key

        values = {
            "중복키": key, "출처": row.get("출처"), "제목": row.get("제목"),
            "발간년도": row.get("발간년도"), "발간등록번호": row["발간등록번호"],
            "URL": row.get("URL"), "라이선스": row.get("라이선스") or config.DEFAULT_LICENSE,
            "첨부파일명": row.get("첨부파일명"), "원문ID": row.get("원문ID"),
            "작성일": row.get("작성일"), "발행기관": row.get("발행기관"), "분류": row.get("분류"),
            "원문페이지": row.get("원문페이지"), "파일경로": row.get("파일경로"),
            "텍스트경로": row.get("텍스트경로"), "추출방법": method,
            "글자수": len(body), "내용해시": digest, "전문": full_text, "수정일시": now,
        }
        cols = ", ".join(values)
        marks = ", ".join("?" * len(values))
        updates = ", ".join(f"{c}=excluded.{c}" for c in values if c not in ("중복키", "라이선스"))
        # 라이선스: DB에 사람이 고친 값이 있으면 기본값으로 덮어쓰지 않음
        updates += (f", 라이선스=CASE WHEN documents.라이선스 IS NULL OR documents.라이선스='' "
                    f"OR documents.라이선스='{config.DEFAULT_LICENSE}' THEN excluded.라이선스 "
                    f"ELSE documents.라이선스 END")
        conn.execute(
            f"INSERT INTO documents ({cols}) VALUES ({marks}) "
            f"ON CONFLICT(중복키) DO UPDATE SET {updates}",
            list(values.values()),
        )
        doc_id = conn.execute("SELECT doc_id FROM documents WHERE 중복키=?", (key,)).fetchone()[0]

        # 조각(chunks) 새로 만들기
        conn.execute("DELETE FROM chunks WHERE doc_id=?", (doc_id,))
        if full_text:
            conn.executemany(
                "INSERT INTO chunks (doc_id, 순번, 페이지, 내용) VALUES (?, ?, ?, ?)",
                [(doc_id, i, page, text) for i, (page, text) in enumerate(split_chunks(full_text))],
            )
        stats["추가/갱신"] += 1

    if has_fts:
        conn.execute("INSERT INTO chunks_fts(chunks_fts) VALUES('rebuild')")
    conn.commit()

    n_docs = conn.execute("SELECT COUNT(*) FROM documents").fetchone()[0]
    n_chunks = conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]

    # 엑셀로도 내보내기 (전문은 너무 길어서 제외)
    summary = pd.read_sql_query("SELECT * FROM documents", conn).drop(columns=["전문"])
    conn.close()
    xlsx = config.DATA_DIR / "최종목록.xlsx"
    try:
        summary.to_excel(xlsx, index=False)
    except PermissionError:
        print(f"! {xlsx.name} 이 Excel에서 열려 있어 저장하지 못했습니다.")

    print(f"처리 결과: {stats}")
    print(f"DB 안의 자료 {n_docs}건, RAG 조각 {n_chunks}개 → {config.DB_PATH}")
    print(f"엑셀 요약 → {xlsx}")


if __name__ == "__main__":
    main()
