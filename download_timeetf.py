"""
timeetf.co.kr 다중 ETF 구성종목 일괄 다운로드.
- 각 ETF는 data/<key>/ 하위 폴더에 저장 (구성종목_YYYY-MM-DD.xlsx)
- 각 ETF의 상장일(최초 실데이터일)부터 오늘까지, 평일(월~금)만
- 표준 라이브러리만 사용 (별도 설치 불필요)

서버는 상장 전·휴장일에도 status 200 + xlsx를 돌려주지만,
그 파일은 헤더 1행짜리 '빈 껍데기'(약 6,460B)다. 시트의 데이터 행 수로
실데이터 여부를 판별해 빈 파일은 저장하지 않는다.
"""

import io
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import zipfile
from datetime import date, timedelta

# start = 상장일(최초 실데이터일). 서버 질의로 확인한 값.
ETFS = [
    {"idx": 6,  "key": "ai_active",        "name": "TIME 글로벌AI인공지능액티브", "start": "2023-05-15"},
    {"idx": 11, "key": "kospi_active",     "name": "TIME 코스피액티브",          "start": "2021-05-24"},
    {"idx": 2,  "key": "nasdaq100_active", "name": "TIME 미국나스닥100액티브",   "start": "2022-05-09"},
    {"idx": 5,  "key": "sp500_active",     "name": "TIME 미국S&P500액티브",      "start": "2022-05-09"},
]

HERE          = os.path.dirname(os.path.abspath(__file__))
ROOT          = os.path.join(HERE, "data")
BASE_URL      = "https://timeetf.co.kr/pdf_excel.php"
END_DATE      = date.today()
SLEEP_SEC     = 0.5
TIMEOUT       = 30
RETRIES       = 3       # 전송 오류/예상치 못한 응답 시 재시도 횟수
REAL_MIN_SIZE = 7000    # 이 크기 이상이면 실데이터로 신뢰 (빈 껍데기는 ~6,460B)

_ROW_RE = re.compile(rb"<row ")


def weekdays(start: date, end: date):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


def is_real_xlsx(data: bytes) -> bool:
    """sheet1 에 헤더 외 데이터 행이 1개 이상 있으면 실데이터로 본다."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            xml = z.read("xl/worksheets/sheet1.xml")
    except Exception:
        return False
    return len(_ROW_RE.findall(xml)) >= 2


def fetch(idx: int, date_str: str) -> tuple[bytes, str, int]:
    qs = urllib.parse.urlencode({"idx": idx, "cate": "", "pdfDate": date_str})
    req = urllib.request.Request(
        f"{BASE_URL}?{qs}",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        return r.read(), r.headers.get("Content-Type", ""), r.status


def fetch_xlsx(idx: int, date_str: str) -> bytes:
    """xlsx 본문을 재시도와 함께 받는다. 끝내 실패하면 예외."""
    last = ""
    for attempt in range(1, RETRIES + 1):
        try:
            data, ctype, status = fetch(idx, date_str)
            if status == 200 and "spreadsheetml" in ctype and data:
                return data
            last = f"status={status} type={ctype} size={len(data)}"
        except Exception as e:  # noqa: BLE001 - 네트워크 예외 전부 재시도 대상
            last = repr(e)
        if attempt < RETRIES:
            time.sleep(SLEEP_SEC * attempt * 2)
    raise RuntimeError(last)


def existing_is_real(path: str) -> bool:
    """이미 받아둔 파일이 실데이터인지. 큰 파일은 크기로 신뢰, 작은 파일만 열어 확인."""
    sz = os.path.getsize(path)
    if sz >= REAL_MIN_SIZE:
        return True
    if sz == 0:
        return False
    with open(path, "rb") as f:
        return is_real_xlsx(f.read())


def run_one(etf: dict) -> tuple[int, int, int, list[str]]:
    out_dir = os.path.join(ROOT, etf["key"])
    os.makedirs(out_dir, exist_ok=True)
    start = date.fromisoformat(etf["start"])
    targets = list(weekdays(start, END_DATE))
    print(f"\n=== [idx={etf['idx']:>2}] {etf['name']} "
          f"({etf['start']} ~ {END_DATE}, {len(targets)} weekdays) ===", flush=True)

    saved = skipped = no_data = removed = 0
    failed: list[str] = []
    for d in targets:
        ds = d.strftime("%Y-%m-%d")
        out = os.path.join(out_dir, f"구성종목_{ds}.xlsx")

        if os.path.exists(out):
            if existing_is_real(out):
                skipped += 1
                continue
            # 빈 껍데기(과거 휴장일 등)가 저장돼 있으면 제거 후 다시 판단
            os.remove(out)
            removed += 1

        try:
            data = fetch_xlsx(etf["idx"], ds)
        except Exception as e:  # noqa: BLE001
            failed.append(ds)
            print(f"  [err]  {ds}: {e}", flush=True)
            time.sleep(SLEEP_SEC)
            continue

        if is_real_xlsx(data):
            with open(out, "wb") as f:
                f.write(data)
            saved += 1
            print(f"  [ok]   {ds}: {len(data):,}B", flush=True)
        else:
            # 휴장일 등 데이터 없는 날 — 정상이므로 실패가 아님
            no_data += 1
        time.sleep(SLEEP_SEC)

    print(f"  -> saved {saved} | skipped {skipped} | no-data {no_data} "
          f"| removed-empty {removed} | failed {len(failed)}", flush=True)
    return saved, skipped, no_data, failed


def main() -> int:
    grand_saved = grand_skipped = grand_no_data = 0
    grand_failed: list[tuple[int, str]] = []
    for etf in ETFS:
        s, k, nd, fails = run_one(etf)
        grand_saved += s
        grand_skipped += k
        grand_no_data += nd
        grand_failed.extend((etf["idx"], d) for d in fails)

    print()
    print(f"[TOTAL] saved {grand_saved} | skipped {grand_skipped} "
          f"| no-data {grand_no_data} | failed {len(grand_failed)}")
    if grand_failed:
        print("Failed (네트워크 오류 등 — 재실행 시 자동 재시도):")
        for idx, ds in grand_failed[:20]:
            print(f"  - idx={idx}  {ds}")
        if len(grand_failed) > 20:
            print(f"  ... {len(grand_failed) - 20} more")
    return 0


if __name__ == "__main__":
    sys.exit(main())
