"""
timeetf.co.kr 다중 ETF 구성종목 일괄 다운로드.
- 각 ETF는 data/<key>/ 하위 폴더에 저장
- 평일(월~금)만, 각 ETF의 상장일부터 (listed — timeetf.co.kr m11_view.php 공시 기준)
- 표준 라이브러리만 사용 (별도 설치 불필요)
- 일시 오류(네트워크/5xx)는 점증 대기 재시도, 저장 전 xlsx(zip) 무결성 검증
- 최근 REFRESH_DAYS 일은 파일이 있어도 재다운로드해 정정 공시 반영
  (내용이 같으면 파일을 건드리지 않아 build_data.py 파싱 캐시가 유지됨)
"""

import io
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from datetime import date, timedelta

ETFS = [
    {"idx": 6,  "key": "ai_active",        "name": "TIME 글로벌AI인공지능액티브", "listed": date(2023, 5, 16)},
    {"idx": 11, "key": "kospi_active",     "name": "TIME 코스피액티브",           "listed": date(2021, 5, 25)},
    {"idx": 2,  "key": "nasdaq100_active", "name": "TIME 미국나스닥100액티브",    "listed": date(2022, 5, 11)},
    {"idx": 5,  "key": "sp500_active",     "name": "TIME 미국S&P500액티브",       "listed": date(2022, 5, 11)},
]

HERE         = os.path.dirname(os.path.abspath(__file__))
ROOT         = os.path.join(HERE, "data")
BASE_URL     = "https://timeetf.co.kr/pdf_excel.php"
END_DATE     = date.today()
SLEEP_SEC    = 0.5
TIMEOUT      = 30
RETRIES      = 3        # 총 시도 횟수 (일시 오류만 재시도)
RETRY_WAIT   = 1.5      # 첫 재시도 대기(초), 이후 ×2 점증
REFRESH_DAYS = 7        # 최근 N일(달력일)은 기존 파일이 있어도 재다운로드 (정정 공시)


def weekdays(start: date, end: date):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


def valid_xlsx(data: bytes) -> bool:
    """저장 전 무결성 검증: zip 컨테이너가 온전하고 워크북이 들어있는지 (stdlib만 사용)."""
    if len(data) < 100 or not data.startswith(b"PK"):
        return False
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            if z.testzip() is not None:
                return False
            names = z.namelist()
            return "[Content_Types].xml" in names and any(n.startswith("xl/") for n in names)
    except (zipfile.BadZipFile, OSError):
        return False


def xlsx_fingerprint(data: bytes):
    """데이터 영역(xl/**)의 (이름, CRC) 지문. 서버가 매 요청마다 xlsx 를 재생성해
    docProps 의 생성 시각 등 메타데이터가 달라지므로, raw 바이트 비교 대신
    실제 시트 데이터가 같은지로 '정정 공시' 여부를 판별한다."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            return sorted((i.filename, i.CRC) for i in z.infolist()
                          if i.filename.startswith("xl/"))
    except (zipfile.BadZipFile, OSError):
        return None


def fetch(idx: int, date_str: str) -> tuple[bytes, str, int]:
    qs = urllib.parse.urlencode({"idx": idx, "cate": "", "pdfDate": date_str})
    req = urllib.request.Request(
        f"{BASE_URL}?{qs}",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        return r.read(), r.headers.get("Content-Type", ""), r.status


def fetch_retry(idx: int, date_str: str):
    """일시 오류(네트워크 단절·타임아웃·5xx)만 재시도. 4xx·정상 응답은 즉시 반환.
    반환: (result | None, last_error | None)"""
    wait = RETRY_WAIT
    last_err = None
    for attempt in range(RETRIES):
        try:
            return fetch(idx, date_str), None
        except urllib.error.HTTPError as e:
            if e.code < 500:
                return None, e          # 4xx → 재시도 무의미
            last_err = e
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            last_err = e
        if attempt < RETRIES - 1:
            time.sleep(wait)
            wait *= 2
    return None, last_err


def save_atomic(path: str, data: bytes) -> None:
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, path)


def run_one(etf: dict) -> tuple[int, int, int, list[str]]:
    out_dir = os.path.join(ROOT, etf["key"])
    os.makedirs(out_dir, exist_ok=True)
    targets = list(weekdays(etf["listed"], END_DATE))
    print(f"\n=== [idx={etf['idx']:>2}] {etf['name']} (상장 {etf['listed']} ~, {len(targets)} weekdays) ===", flush=True)
    saved = updated = skipped = 0
    failed: list[str] = []
    for d in targets:
        ds = d.strftime("%Y-%m-%d")
        out = os.path.join(out_dir, f"구성종목_{ds}.xlsx")
        exists = os.path.exists(out) and os.path.getsize(out) > 0
        recent = (END_DATE - d).days < REFRESH_DAYS
        if exists and not recent:
            skipped += 1
            continue

        res, err = fetch_retry(etf["idx"], ds)
        if res is None:
            if exists:
                print(f"  [keep] {ds}: 재확인 실패({err}) → 기존 파일 유지", flush=True)
            else:
                failed.append(ds)
                print(f"  [err]  {ds}: {err}", flush=True)
            time.sleep(SLEEP_SEC)
            continue

        data, ctype, status = res
        ok = status == 200 and "spreadsheetml" in ctype and valid_xlsx(data)
        if not ok:
            if exists:
                print(f"  [keep] {ds}: 응답 이상(status={status}, type={ctype}, "
                      f"size={len(data)}) → 기존 파일 유지", flush=True)
            else:
                failed.append(ds)
                print(f"  [fail] {ds}: status={status} type={ctype} size={len(data)}"
                      f"{' (zip 무결성 불합격)' if status == 200 and 'spreadsheetml' in ctype else ''}",
                      flush=True)
            time.sleep(SLEEP_SEC)
            continue

        if exists:
            with open(out, "rb") as f:
                old = f.read()
            if old == data or xlsx_fingerprint(old) == xlsx_fingerprint(data):
                skipped += 1                      # 시트 데이터 동일 → mtime 보존(파싱 캐시 유지)
            else:
                save_atomic(out, data)
                updated += 1
                print(f"  [upd]  {ds}: 정정 공시 반영 {len(old):,}B → {len(data):,}B", flush=True)
        else:
            save_atomic(out, data)
            saved += 1
            print(f"  [ok]   {ds}: {len(data):,}B", flush=True)
        time.sleep(SLEEP_SEC)
    print(f"  -> saved {saved} | updated {updated} | skipped {skipped} | failed {len(failed)}", flush=True)
    return saved, updated, skipped, failed


def main() -> int:
    grand_saved = grand_updated = grand_skipped = 0
    grand_failed: list[tuple[int, str]] = []
    for etf in ETFS:
        s, u, k, fails = run_one(etf)
        grand_saved += s
        grand_updated += u
        grand_skipped += k
        grand_failed.extend((etf["idx"], d) for d in fails)
    print()
    print(f"[TOTAL] saved {grand_saved} | updated {grand_updated} | "
          f"skipped {grand_skipped} | failed {len(grand_failed)}")
    if grand_failed:
        print("Failed (휴장/데이터 없음 가능):")
        for idx, ds in grand_failed[:20]:
            print(f"  - idx={idx}  {ds}")
        if len(grand_failed) > 20:
            print(f"  ... {len(grand_failed) - 20} more")
    return 0


if __name__ == "__main__":
    sys.exit(main())
