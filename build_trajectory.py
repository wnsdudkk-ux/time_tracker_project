"""
data.js(보유 시계열) + prices.json(yfinance 현지통화 수정주가) →
trajectory.js (window.TRAJ_DATA) 빌드.

각 ETF의 영업일 축(data.js 의 dates)에 종목별 현지 주가를 직전 거래일 기준
forward-fill 로 정렬하고, 종목의 '보유 구간'(첫 보유일~마지막 보유일)만 채운다.
평균단가·손익·능동매매 등 파생지표는 브라우저(JS)에서 계산한다 — 여기서는 raw 주가만.
"""

import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_JS = os.path.join(HERE, "data.js")
PRICES = os.path.join(HERE, "prices.json")
OUT_JS = os.path.join(HERE, "trajectory.js")


def load_data_js():
    s = open(DATA_JS, encoding="utf-8").read()
    return json.loads(s[s.index("{"): s.rindex("}") + 1])


def align(fund_dates, p_dates, p_close, lo, hi):
    """fund_dates[lo..hi] 각 날짜에 대해 그 날짜 이하의 가장 최근 종가(forward-fill).
       p_dates 는 오름차순. 보유구간 밖(lo 미만, hi 초과)은 None."""
    out = [None] * len(fund_dates)
    j = 0
    n = len(p_dates)
    last = None
    # lo 이전 가격도 소비해서 lo 시점의 forward-fill 기준을 맞춘다
    for i in range(0, hi + 1):
        d = fund_dates[i]
        while j < n and p_dates[j] <= d:
            last = p_close[j]
            j += 1
        if i >= lo:
            out[i] = last
    return out


def build_etf(etf, prices):
    dates = etf["dates"]
    series = etf["series"]
    px, ccy = {}, {}
    for t in etf["tickers"]:
        if t == "__CASH__" or t not in prices:
            continue
        w = series[t]["weight"]
        held = [i for i, x in enumerate(w) if x is not None]
        if not held:
            continue
        lo, hi = held[0], held[-1]
        rec = prices[t]
        aligned = align(dates, rec["dates"], rec["close"], lo, hi)
        if not any(v is not None for v in aligned):
            continue
        px[t] = aligned
        ccy[t] = rec["ccy"]
    return {"idx": etf["idx"], "name": etf["name"], "px": px, "ccy": ccy}


def main():
    if not os.path.exists(PRICES):
        print(f"[!] {PRICES} 없음 → 먼저 fetch_prices.py 실행")
        return 1
    d = load_data_js()
    prices = json.load(open(PRICES, encoding="utf-8"))
    out = {"generated_at": d.get("generated_at", ""), "etfs": []}
    for etf in d["etfs"]:
        e = build_etf(etf, prices)
        out["etfs"].append(e)
        n_px = len(e["px"])
        print(f"  - {etf['name']}: 가격연계 {n_px}/{len(etf['tickers'])-1}종목")
    with open(OUT_JS, "w", encoding="utf-8") as f:
        f.write("window.TRAJ_DATA = ")
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")
    print(f"\n[OK] {OUT_JS}  ({os.path.getsize(OUT_JS)/1e6:.2f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
