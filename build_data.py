"""
모든 ETF의 xlsx → 단일 data.js (window.ETF_DATA = {generated_at, etfs:[...]}) 빌드.
표준 라이브러리 + openpyxl 만 사용.
"""

import json
import os
import re
import sys
from datetime import datetime, date as Date, timedelta
from glob import glob

import openpyxl

ETFS = [
    {"idx": 6,  "key": "ai_active",        "name": "TIME 글로벌AI인공지능액티브"},
    {"idx": 11, "key": "kospi_active",     "name": "TIME 코스피액티브"},
    {"idx": 2,  "key": "nasdaq100_active", "name": "TIME 미국나스닥100액티브"},
    {"idx": 5,  "key": "sp500_active",     "name": "TIME 미국S&P500액티브"},
]

HERE     = os.path.dirname(os.path.abspath(__file__))
ROOT     = os.path.join(HERE, "data")
OUT_JS   = os.path.join(HERE, "data.js")
KEYS_TXT = os.path.join(HERE, "api_keys.txt")
KEYS_JS  = os.path.join(HERE, "keys.js")

DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")


def parse_number(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.replace(",", "").strip()
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def normalize_ticker(code, name):
    if (code is None or code == "") and isinstance(name, str) and name.strip() == "현금":
        return "__CASH__"
    if isinstance(code, (int, float)):
        return f"{int(code):06d}"
    if isinstance(code, str):
        return code.strip()
    return str(code)


def classify_region(ticker, name):
    if ticker == "__CASH__":
        return "Cash"
    if " US EQUITY" in ticker:
        return "US"
    if " HK EQUITY" in ticker:
        return "HK"
    if " JP EQUITY" in ticker:
        return "JP"
    if " CH EQUITY" in ticker or " C2 EQUITY" in ticker:
        return "CN"
    if re.fullmatch(r"\d{6}", ticker):
        return "KR"
    if re.fullmatch(r"[0-9A-Z]{6}", ticker):
        return "Fund"
    return "Other"


def read_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    out = []
    rows_iter = ws.iter_rows(values_only=True)
    next(rows_iter, None)  # 헤더 skip
    for row in rows_iter:
        if row is None or len(row) < 5:
            continue
        code, name, qty, value, weight = row[0], row[1], row[2], row[3], row[4]
        if name is None:
            continue
        ticker = normalize_ticker(code, name)
        out.append({
            "ticker": ticker,
            "name": str(name).strip(),
            "qty": parse_number(qty),
            "value": parse_number(value),
            "weight": parse_number(weight),
        })
    wb.close()
    return out


def build_etf(etf: dict):
    xlsx_dir = os.path.join(ROOT, etf["key"])
    if not os.path.isdir(xlsx_dir):
        print(f"[!] {etf['name']}: 폴더 없음 → 스킵 ({xlsx_dir})")
        return None
    pattern = os.path.join(xlsx_dir, "구성종목_*.xlsx")
    files = [f for f in sorted(glob(pattern)) if not os.path.basename(f).startswith("~$")]
    if not files:
        print(f"[!] {etf['name']}: xlsx 없음 → 스킵")
        return None

    print(f"[*] {etf['name']}: {len(files)}개 파일")
    by_date = {}
    for f in files:
        m = DATE_RE.search(os.path.basename(f))
        if not m:
            continue
        ds = m.group(1)
        try:
            holdings = read_xlsx(f)
            if holdings:
                by_date[ds] = holdings
        except Exception as e:
            print(f"    ! {ds}: {e}")

    dates = sorted(by_date.keys())
    if not dates:
        return None

    tickers = {}
    for ds in dates:
        for h in by_date[ds]:
            tickers[h["ticker"]] = {
                "name": h["name"],
                "region": classify_region(h["ticker"], h["name"]),
            }

    n = len(dates)
    series = {
        t: {"weight": [None] * n, "qty": [None] * n, "value": [None] * n}
        for t in tickers
    }
    weight_sums = []
    for di, ds in enumerate(dates):
        s = 0.0
        for h in by_date[ds]:
            t = h["ticker"]
            series[t]["weight"][di] = h["weight"]
            series[t]["qty"][di] = h["qty"]
            series[t]["value"][di] = h["value"]
            if h["weight"] is not None:
                s += h["weight"]
        weight_sums.append(round(s, 2))

    d_start = Date.fromisoformat(dates[0])
    d_end = Date.fromisoformat(dates[-1])
    missing = []
    cur = d_start
    while cur <= d_end:
        if cur.weekday() < 5:
            ds = cur.isoformat()
            if ds not in by_date:
                missing.append(ds)
        cur += timedelta(days=1)

    invalid = [[d, s] for d, s in zip(dates, weight_sums) if abs(s - 100) > 0.5]

    return {
        "idx": etf["idx"],
        "name": etf["name"],
        "meta": {
            "first_date": dates[0],
            "last_date": dates[-1],
            "n_dates": n,
            "n_tickers": len(tickers),
        },
        "dates": dates,
        "tickers": tickers,
        "series": series,
        "weight_sums": weight_sums,
        "missing_dates": missing,
        "invalid_dates": invalid,
    }


def write_keys_js() -> None:
    """api_keys.txt → keys.js (window.API_KEYS = {...}). 파일 없으면 skip."""
    if not os.path.exists(KEYS_TXT):
        if os.path.exists(KEYS_JS):
            os.remove(KEYS_JS)
        print(f"[!] {KEYS_TXT} 없음 → keys.js 미생성 (LLM 보고서 비활성)")
        return
    keys: dict[str, str] = {}
    with open(KEYS_TXT, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                k, v = line.split("=", 1)
                keys[k.strip()] = v.strip()
    with open(KEYS_JS, "w", encoding="utf-8") as f:
        f.write("window.API_KEYS = ")
        json.dump(keys, f, ensure_ascii=False)
        f.write(";\n")
    print(f"[OK] {KEYS_JS}  (키 {len(keys)}개)")


def main() -> int:
    out = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "etfs": [],
    }
    for etf in ETFS:
        d = build_etf(etf)
        if d:
            out["etfs"].append(d)
    if not out["etfs"]:
        print("[!] 빌드된 ETF 없음. 먼저 다운로드를 실행하세요.")
        return 1

    with open(OUT_JS, "w", encoding="utf-8") as f:
        f.write("window.ETF_DATA = ")
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")

    size_mb = os.path.getsize(OUT_JS) / (1024 * 1024)
    print()
    print(f"[OK] {OUT_JS}  ({size_mb:.2f} MB)")
    write_keys_js()
    for e in out["etfs"]:
        m = e["meta"]
        print(f"   - {e['name']}: {m['n_dates']}일, {m['n_tickers']}종목"
              f"  (누락 {len(e['missing_dates'])}일, 비중합 이상 {len(e['invalid_dates'])}일)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
