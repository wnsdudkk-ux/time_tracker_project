"""
timeetf.co.kr 다중 ETF 구성종목 일괄 다운로드.
- 각 ETF는 timeetf_xlsx/idx_<N>/ 하위 폴더에 저장
- 평일(월~금)만, 최근 1년
- 표준 라이브러리만 사용 (별도 설치 불필요)
"""

import os
import sys
import time
import urllib.parse
import urllib.request
from datetime import date, timedelta

ETFS = [
    {"idx": 6,  "key": "ai_active",        "name": "TIME 글로벌AI인공지능액티브"},
    {"idx": 11, "key": "kospi_active",     "name": "TIME 코스피액티브"},
    {"idx": 2,  "key": "nasdaq100_active", "name": "TIME 미국나스닥100액티브"},
    {"idx": 5,  "key": "sp500_active",     "name": "TIME 미국S&P500액티브"},
]

HERE       = os.path.dirname(os.path.abspath(__file__))
ROOT       = os.path.join(HERE, "data")
BASE_URL   = "https://timeetf.co.kr/pdf_excel.php"
END_DATE   = date.today()
START_DATE = END_DATE - timedelta(days=365)
SLEEP_SEC  = 0.5
TIMEOUT    = 30


def weekdays(start: date, end: date):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


def fetch(idx: int, date_str: str) -> tuple[bytes, str, int]:
    qs = urllib.parse.urlencode({"idx": idx, "cate": "", "pdfDate": date_str})
    req = urllib.request.Request(
        f"{BASE_URL}?{qs}",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        return r.read(), r.headers.get("Content-Type", ""), r.status


def run_one(etf: dict) -> tuple[int, int, list[str]]:
    out_dir = os.path.join(ROOT, etf["key"])
    os.makedirs(out_dir, exist_ok=True)
    targets = list(weekdays(START_DATE, END_DATE))
    print(f"\n=== [idx={etf['idx']:>2}] {etf['name']} ({len(targets)} weekdays) ===", flush=True)
    saved = skipped = 0
    failed: list[str] = []
    for d in targets:
        ds = d.strftime("%Y-%m-%d")
        out = os.path.join(out_dir, f"구성종목_{ds}.xlsx")
        if os.path.exists(out) and os.path.getsize(out) > 0:
            skipped += 1
            continue
        try:
            data, ctype, status = fetch(etf["idx"], ds)
        except Exception as e:
            failed.append(ds)
            print(f"  [err]  {ds}: {e}", flush=True)
            time.sleep(SLEEP_SEC)
            continue
        if status == 200 and "spreadsheetml" in ctype and len(data) > 0:
            with open(out, "wb") as f:
                f.write(data)
            saved += 1
            print(f"  [ok]   {ds}: {len(data):,}B", flush=True)
        else:
            failed.append(ds)
            print(f"  [fail] {ds}: status={status} type={ctype} size={len(data)}", flush=True)
        time.sleep(SLEEP_SEC)
    print(f"  -> saved {saved} | skipped(existing) {skipped} | failed {len(failed)}", flush=True)
    return saved, skipped, failed


def main() -> int:
    grand_saved = 0
    grand_skipped = 0
    grand_failed: list[tuple[int, str]] = []
    for etf in ETFS:
        s, k, fails = run_one(etf)
        grand_saved += s
        grand_skipped += k
        grand_failed.extend((etf["idx"], d) for d in fails)
    print()
    print(f"[TOTAL] saved {grand_saved} | skipped {grand_skipped} | failed {len(grand_failed)}")
    if grand_failed:
        print("Failed (휴장/데이터 없음 가능):")
        for idx, ds in grand_failed[:20]:
            print(f"  - idx={idx}  {ds}")
        if len(grand_failed) > 20:
            print(f"  ... {len(grand_failed) - 20} more")
    return 0


if __name__ == "__main__":
    sys.exit(main())
