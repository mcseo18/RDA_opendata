# -*- coding: utf-8 -*-
"""
1단계: 공공데이터포털 API로 '자료 목록'을 모아 Excel/CSV로 저장합니다.

  수집 대상 (config.py 의 APIS 에서 바꿀 수 있습니다)
    - 국립원예특작과학원 연구동향 조회 서비스   → 첨부파일 다운로드 URL 수집
    - 국립원예특작과학원 원예·특용작물 기술정보 → 본문(contentsKor)과 첨부파일 URL 수집

  저장 결과 (data/lists 폴더)
    - 목록_연구동향.xlsx / .csv
    - 목록_기술정보.xlsx / .csv
    - 목록_전체.csv   ← 2단계(다운로드)와 4단계(DB)가 읽는 파일

  실행 방법 (명령 프롬프트나 PowerShell에서, 이 폴더로 이동한 뒤)
    python step1_collect_list.py            → 두 API 모두 전체 수집
    python step1_collect_list.py --test     → 1페이지만 받아서 어떤 칸이 오는지 보여줌 (처음에 꼭 해보세요)
    python step1_collect_list.py --only 기술정보   → 한 API만 수집
"""

import argparse
import html
import re
import sys
import time
import xml.etree.ElementTree as ET

import pandas as pd
import requests

import config

# 결과 표의 칸(열) 순서. 앞의 6개가 요청하신 기본 칸입니다.
COLUMNS = [
    "출처", "제목", "발간년도", "발간등록번호", "URL", "라이선스",
    "첨부파일명", "원문ID", "작성일", "발행기관", "분류", "원문페이지", "본문",
]

# API마다 같은 뜻의 칸 이름이 조금씩 다를 수 있어서, 후보 이름을 여러 개 둡니다.
# (앞에 있는 이름부터 찾아봅니다)
TITLE_KEYS = ["subject", "title", "ttl", "sj", "nttSj", "bbsSj", "titl"]
DATE_KEYS = ["regDate", "regDt", "registDt", "wrtDate", "writeDate", "wrtDt",
             "pblicteDe", "pubDate", "date", "rgsde"]
ID_KEYS = ["contentId", "idx", "nttId", "seq", "id", "no"]
BODY_KEYS = ["contentsKor", "contents", "content", "cn", "nttCn"]
ORG_KEYS = ["publishOrg", "orgNm", "writer", "wrter", "author", "register"]
CATEGORY_KEYS = ["repcategory", "category", "ctgry", "ftrSeCode"]
PAGE_URL_KEYS = ["originUrl", "linkUrl", "pageUrl", "link"]

# 발간등록번호 모양: 11-1390000-000000-01 (숫자2-숫자7-숫자6-숫자2)
PUB_NO_PATTERN = re.compile(r"\b\d{2}-\d{7}-\d{6}-\d{2}\b")

# 공공데이터포털 공통 오류코드 → 쉬운 설명
ERROR_HELP = {
    "1": "API 서버 내부 오류입니다. 잠시 뒤 다시 시도하세요.",
    "10": "요청 값이 잘못되었습니다. config.py 의 파라미터를 확인하세요.",
    "12": "API 주소가 없거나 폐기되었습니다. data.go.kr 상세페이지에서 주소를 다시 확인하세요.",
    "20": "이 API에 대한 활용신청이 승인되지 않았습니다. data.go.kr 마이페이지에서 확인하세요.",
    "22": "하루 호출 한도를 넘었습니다. 내일 다시 실행하세요.",
    "30": "등록되지 않은 인증키입니다. 키를 정확히 붙여넣었는지, 승인 후 1~2시간 지났는지 확인하세요.",
    "31": "인증키 사용 기간이 끝났습니다. 연장 신청하세요.",
    "32": "등록되지 않은 IP입니다.",
}


# ─────────────────────────────────────────────
# 작은 도우미 함수들
# ─────────────────────────────────────────────
def pick(record, keys):
    """record(딕셔너리)에서 keys 후보 중 처음으로 값이 있는 것을 돌려줍니다."""
    lower_map = {k.lower(): v for k, v in record.items()}
    for key in keys:
        value = lower_map.get(key.lower())
        if value:
            return value
    return ""


def clean_text(text):
    """HTML 태그와 &nbsp; 같은 기호를 없애고 공백을 정리합니다."""
    if not text:
        return ""
    text = re.sub(r"<br\s*/?>|</p>|</div>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    text = re.sub(r"[ \t ]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def find_year(text):
    """날짜 문자열에서 4자리 연도(1990~2099)를 찾습니다."""
    match = re.search(r"(19[9]\d|20\d\d)", text or "")
    return match.group(1) if match else ""


def find_pub_no(*texts):
    """제목/본문/파일명 등에서 발간등록번호를 찾아봅니다. (없으면 빈칸)"""
    for text in texts:
        match = PUB_NO_PATTERN.search(text or "")
        if match:
            return match.group(0)
    return ""


# ─────────────────────────────────────────────
# API 호출과 XML 해석
# ─────────────────────────────────────────────
def call_api(url, params):
    """API를 호출해서 XML 루트를 돌려줍니다. 실패하면 몇 번 더 시도합니다."""
    headers = {"User-Agent": config.USER_AGENT}
    last_error = None
    for attempt in range(1, config.RETRY_COUNT + 1):
        try:
            resp = requests.get(url, params=params, headers=headers,
                                timeout=config.REQUEST_TIMEOUT)
            # 인증키 오류일 때 서버가 403 과 함께 오류 XML을 주므로, 먼저 XML로 읽어봅니다
            try:
                return ET.fromstring(resp.content)
            except ET.ParseError:
                resp.raise_for_status()
                # XML이 아닌 글자(예: 인증 오류 문구)가 오면 앞부분을 보여줍니다
                raise RuntimeError(f"XML이 아닌 응답: {resp.text[:200]}")
        except (requests.RequestException, RuntimeError) as err:
            last_error = err
            print(f"  ! 호출 실패({attempt}/{config.RETRY_COUNT}): {err}")
            time.sleep(config.API_DELAY_SEC * attempt * 2)
    raise RuntimeError(f"API 호출을 계속 실패했습니다: {last_error}")


def check_error(root):
    """공공데이터포털 공통 오류 응답이면 설명을 출력하고 프로그램을 멈춥니다."""
    code = root.findtext(".//returnReasonCode")
    if code and code.strip() not in ("0", "00"):
        code = code.strip()
        msg = root.findtext(".//returnAuthMsg") or root.findtext(".//errMsg") or ""
        print(f"\n[API 오류 {code}] {msg}")
        print("  → " + ERROR_HELP.get(code, "data.go.kr 오류코드 표를 확인하세요."))
        sys.exit(1)


def leaves(element):
    """element 아래의 '끝 칸'(자식이 없는 태그)들을 {태그이름: 값} 딕셔너리로 만듭니다."""
    record = {}
    for child in element.iter():
        if child is element or len(child):
            continue
        tag = child.tag.split("}")[-1]           # XML 네임스페이스 제거
        value = (child.text or "").strip()
        if tag in record and value and not record[tag]:
            record[tag] = value
        else:
            record.setdefault(tag, value)
    return record


def parse_items(root):
    """
    XML 안에서 '자료 1건'에 해당하는 묶음들을 찾아 딕셔너리 목록으로 돌려줍니다.
    보통은 <item> 태그지만, 없으면 가장 많이 반복되는 묶음 태그를 대신 씁니다.
    """
    items = root.findall(".//item")
    if not items:
        skip = {"header", "cmmmsgheader", "response", "body", "items"}
        counts = {}
        for el in root.iter():
            tag = el.tag.split("}")[-1]
            if tag.lower() in skip:
                continue
            leaf_children = [c for c in el if not len(c)]
            if len(leaf_children) >= 3:
                counts.setdefault(tag, []).append(el)
        if counts:
            items = max(counts.values(), key=len)
    return [leaves(it) for it in items]


def attachments(record):
    """
    한 건의 자료에서 첨부파일 (파일명, URL) 쌍을 모두 찾습니다.
    예) atchfileNm1 + atchfileUrl1, atchfileNm2 + atchfileUrl2 ...
    """
    page_keys = {k.lower() for k in PAGE_URL_KEYS}
    found = []
    for key, value in record.items():
        if not value or key.lower() in page_keys:
            continue
        looks_like_url = value.startswith(("http://", "https://"))
        url_key = re.search(r"(url|down|path|file)", key, re.I)
        if looks_like_url and url_key:
            # 짝이 되는 파일명 칸 찾기: atchfileUrl1 → atchfileNm1
            name = ""
            for alt in ("Nm", "Name", "nm", "name", "_nm", "_name"):
                name_key = re.sub(r"(?i)url", alt, key)
                if record.get(name_key):
                    name = record[name_key]
                    break
            found.append((name, value))
    return found


# ─────────────────────────────────────────────
# 한 API 전체 수집
# ─────────────────────────────────────────────
def collect_api(api_id, test_mode=False):
    api = config.APIS[api_id]
    key = config.get_service_key()
    print(f"\n=== {api['name']} 목록 수집 시작 ===")

    rows = []
    seen_ids = set()
    page = 1
    total_count = None

    while page <= config.MAX_PAGES:
        params = dict(api["extra_params"])
        params = {k: v for k, v in params.items() if v != ""}   # 빈 값은 보내지 않음
        params[api["key_param"]] = key
        params[api["page_param"]] = page

        root = call_api(api["list_url"], params)
        check_error(root)
        items = parse_items(root)

        if test_mode:
            show_test_result(root, items, api)
            return []

        # 전체 건수가 응답에 있으면 기억해 둡니다 (반복 종료 판단용)
        if total_count is None:
            tc = root.findtext(".//totalCount") or root.findtext(".//resultCnt")
            total_count = int(tc) if tc and tc.strip().isdigit() else None

        new_items = []
        for item in items:
            item_id = pick(item, ID_KEYS) or pick(item, TITLE_KEYS)
            if item_id in seen_ids:
                continue
            seen_ids.add(item_id)
            new_items.append(item)

        if not new_items:          # 새 자료가 없으면 마지막 페이지를 지난 것
            break

        for item in new_items:
            # 목록에 첨부파일 정보가 없으면 '상세 조회'를 한 번 더 합니다.
            if not attachments(item) and api.get("use_detail"):
                item = merge_detail(api, key, item)
            rows.extend(item_to_rows(api["name"], item))

        print(f"  {page}페이지: {len(new_items)}건 (누적 자료 {len(seen_ids)}건)")
        if total_count and len(seen_ids) >= total_count:
            break
        page += 1
        time.sleep(config.API_DELAY_SEC)

    df = pd.DataFrame(rows, columns=COLUMNS)
    print(f"=== {api['name']}: 자료 {len(seen_ids)}건, 행 {len(df)}개 ===")
    return df


def merge_detail(api, key, item):
    """상세 조회 API로 빠진 정보(첨부파일 등)를 채웁니다."""
    item_id = pick(item, ID_KEYS)
    if not item_id:
        return item
    time.sleep(config.API_DELAY_SEC)
    try:
        root = call_api(api["detail_url"], {api["key_param"]: key, "idx": item_id})
    except RuntimeError as err:
        print(f"  ! 상세 조회 실패 (idx={item_id}): {err}")
        return item
    found = parse_items(root) or [leaves(root)]
    merged = dict(item)
    for k, v in found[0].items():
        if v and not merged.get(k):
            merged[k] = v
    return merged


def item_to_rows(source_name, item):
    """자료 1건을 결과 표의 행으로 바꿉니다. 첨부파일이 여러 개면 파일마다 한 행씩."""
    title = clean_text(pick(item, TITLE_KEYS))
    date = pick(item, DATE_KEYS)
    body = clean_text(pick(item, BODY_KEYS))
    files = attachments(item)
    # 같은 자료가 PDF와 한글(hwp/hwpx) 두 형식으로 함께 올라온 경우가 많습니다.
    # PDF가 있으면 같은 내용의 한글 파일은 빼서 중복 다운로드를 막습니다.
    if any(name.lower().endswith(".pdf") for name, _ in files):
        files = [(n, u) for n, u in files if not n.lower().endswith((".hwp", ".hwpx"))]
    files = files or [("", "")]   # 첨부가 없어도 본문 때문에 1행은 남김

    rows = []
    for file_name, url in files:
        rows.append({
            "출처": source_name,
            "제목": title,
            "발간년도": find_year(date),
            "발간등록번호": find_pub_no(title, file_name, body[:3000]),
            "URL": url,
            "라이선스": config.DEFAULT_LICENSE,
            "첨부파일명": file_name,
            "원문ID": pick(item, ID_KEYS),
            "작성일": date,
            "발행기관": pick(item, ORG_KEYS),
            "분류": pick(item, CATEGORY_KEYS),
            "원문페이지": pick(item, PAGE_URL_KEYS),
            "본문": body,
        })
    return rows


def show_test_result(root, items, api):
    """--test 모드: 실제 응답 모양을 보여줍니다."""
    print("\n[응답 앞부분]")
    print(ET.tostring(root, encoding="unicode")[:1500])
    print(f"\n[찾은 자료 수] {len(items)}건")
    if items:
        print("\n[첫 번째 자료의 칸 이름과 값]")
        for k, v in items[0].items():
            print(f"  {k:20s} : {v[:80]}")
        print("\n[인식된 첨부파일]", attachments(items[0]) or "없음 (상세 조회에서 찾습니다)")
    print("\n칸 이름이 위 후보 목록(TITLE_KEYS 등)과 다르면 이 파일 위쪽 목록에 추가하세요.")


# ─────────────────────────────────────────────
# 저장
# ─────────────────────────────────────────────
def save(df, file_stem):
    """같은 내용을 CSV(엑셀에서 한글 깨지지 않게 utf-8-sig)와 xlsx로 저장합니다."""
    csv_path = config.LIST_DIR / f"{file_stem}.csv"
    xlsx_path = config.LIST_DIR / f"{file_stem}.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    # Excel 한 칸에는 약 32,000자까지만 들어가므로 본문을 잘라서 저장합니다.
    # (전체 본문은 CSV에 그대로 있습니다)
    df_x = df.copy()
    df_x["본문"] = df_x["본문"].fillna("").str.slice(0, 32000)
    try:
        df_x.to_excel(xlsx_path, index=False)
    except PermissionError:
        print(f"  ! {xlsx_path.name} 이 Excel에서 열려 있어 저장하지 못했습니다. 닫고 다시 실행하세요.")
    print(f"  저장: {csv_path}")


def main():
    parser = argparse.ArgumentParser(description="농진청 자료 목록 수집")
    parser.add_argument("--test", action="store_true", help="1페이지만 받아 응답 모양 확인")
    parser.add_argument("--only", choices=list(config.APIS), help="한 API만 수집")
    args = parser.parse_args()

    config.make_folders()
    targets = [args.only] if args.only else list(config.APIS)

    frames = []
    for api_id in targets:
        df = collect_api(api_id, test_mode=args.test)
        if args.test:
            continue
        save(df, f"목록_{api_id}")
        frames.append(df)

    if args.test:
        print("\n테스트가 끝났습니다. 문제가 없으면 --test 없이 다시 실행하세요.")
        return

    # 이번에 수집한 결과를 기존 '목록_전체.csv'와 합칩니다 (--only 로 나눠 받아도 누적됨)
    all_df = pd.concat(frames, ignore_index=True)
    if config.ALL_LIST_CSV.exists():
        old = pd.read_csv(config.ALL_LIST_CSV, dtype=str, keep_default_na=False)
        old = old[~old["출처"].isin(all_df["출처"].unique())]   # 다시 받은 출처는 새 것으로 교체
        all_df = pd.concat([old, all_df], ignore_index=True)
    all_df.to_csv(config.ALL_LIST_CSV, index=False, encoding="utf-8-sig")
    print(f"\n완료! 전체 목록 {len(all_df)}행 → {config.ALL_LIST_CSV}")
    print("다음 단계: python step2_download_pdf.py")


if __name__ == "__main__":
    main()
