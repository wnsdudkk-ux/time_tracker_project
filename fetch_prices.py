"""
data.js 의 전 종목(현재+과거청산)을 yfinance 심볼로 매핑해 현지통화 수정주가를
일별로 받아 prices.json 으로 저장한다.

- 가격은 auto_adjust=True (액면분할·배당 반영 수정주가). 통화는 거래소 접미사로 추론.
- 종목별 캐시(price_cache/<ticker>.json)로 재실행 시 이어받기 / 빠른 갱신.
- 매핑 불가(Fund, 지수선물, ISIN 등)·데이터 없음은 건너뛰고 로그.

사용:  python fetch_prices.py            # 전체
       python fetch_prices.py --test     # 대표 몇 종목만 점검
       python fetch_prices.py --refresh  # 캐시 무시하고 다시 받기
"""

import json
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_JS = os.path.join(HERE, "data.js")
OUT_JSON = os.path.join(HERE, "prices.json")
CACHE = os.path.join(HERE, "price_cache")
HISTORY_START = "2021-05-01"   # 가장 이른 상장일(코스피 2021-05-24)보다 약간 앞


def load_pairs():
    """data.js → {ticker: region} 합집합, 그리고 오늘 날짜(=fund 마지막일 다음)."""
    s = open(DATA_JS, encoding="utf-8").read()
    s = s[s.index("{"): s.rindex("}") + 1]
    d = json.loads(s)
    pairs, last = {}, ""
    for e in d["etfs"]:
        last = max(last, e["dates"][-1])
        for t, info in e["tickers"].items():
            pairs[t] = info["region"]
    return pairs, last


def candidates(ticker: str, region: str):
    """(yf_symbol, currency) 후보 리스트. KR 은 .KS→.KQ 순으로 시도."""
    t = ticker.strip()
    head = t.split(" ")[0]
    if region == "US":
        return [(head.replace("/", "-"), "USD")]
    if region == "KR":
        return [(f"{head}.KS", "KRW"), (f"{head}.KQ", "KRW")] if head.isdigit() else []
    if region == "HK":
        return [(f"{int(head):04d}.HK", "HKD")] if head.isdigit() else []
    if region == "JP":
        return [(f"{head}.T", "JPY")]      # 285A 같은 영숫자 코드도 .T
    if region == "CN":
        if not head.isdigit():
            return []
        return [(f"{head}.SS" if head.startswith("6") else f"{head}.SZ", "CNY")]
    return []   # Fund / Other(ISIN·지수선물) / Cash → 매핑 불가


def extract_close(df, sym):
    """yfinance DataFrame → (dates[], close[]) 수정종가. 단일 종목 기준."""
    if df is None or len(df) == 0:
        return None
    cl = df["Close"]
    if hasattr(cl, "columns"):          # MultiIndex 컬럼이면 첫 열
        cl = cl.iloc[:, 0]
    cl = cl.dropna()
    if len(cl) == 0:
        return None
    dates = [d.strftime("%Y-%m-%d") for d in cl.index]
    close = [round(float(x), 4) for x in cl.values]
    return dates, close


def fetch_one(yf, cands, end):
    for sym, ccy in cands:
        try:
            df = yf.download(sym, start=HISTORY_START, end=end, auto_adjust=True,
                             progress=False, threads=False, timeout=30)
        except Exception:
            df = None
        got = extract_close(df, sym)
        if got:
            return sym, ccy, got[0], got[1]
    return None


def main():
    args = set(sys.argv[1:])
    test = "--test" in args
    refresh = "--refresh" in args
    os.makedirs(CACHE, exist_ok=True)

    import yfinance as yf
    from datetime import date, timedelta
    pairs, last_fund = load_pairs()
    end = (date.fromisoformat(last_fund) + timedelta(days=2)).isoformat()
    fresh_cutoff = (date.fromisoformat(last_fund) - timedelta(days=4)).isoformat()  # 캐시가 이보다 최신이면 스킵
    if test:
        pick = ["NVDA US EQUITY", "005930", "700 HK EQUITY", "285A JP EQUITY",
                "300274 CH EQUITY", "0043Y0", "NQU3 Index"]
        pairs = {t: pairs.get(t, "KR" if t.isdigit() else "Other") for t in pick}

    items = sorted(pairs.items())
    print(f"대상 {len(items)}종목 · 기간 {HISTORY_START}~{end}", flush=True)

    ok = skip = fail = 0
    skipped, failed = [], []
    for i, (ticker, region) in enumerate(items, 1):
        safe = ticker.replace("/", "_").replace(" ", "_").replace("*", "x")
        cf = os.path.join(CACHE, f"{safe}.json")
        if os.path.exists(cf) and not refresh:
            try:
                last = (json.load(open(cf, encoding="utf-8")).get("dates") or [""])[-1]
            except Exception:
                last = ""
            if last >= fresh_cutoff:     # 캐시가 충분히 최신 → 스킵
                ok += 1
                continue
        cands = candidates(ticker, region)
        if not cands:
            skip += 1
            skipped.append(f"{ticker} ({region})")
            continue
        res = fetch_one(yf, cands, end)
        if res:
            sym, ccy, dates, close = res
            json.dump({"ticker": ticker, "yf": sym, "ccy": ccy,
                       "dates": dates, "close": close},
                      open(cf, "w", encoding="utf-8"))
            ok += 1
            print(f"  [{i}/{len(items)}] {ticker:24s} -> {sym:12s} {ccy} {len(dates)}d", flush=True)
        else:
            fail += 1
            failed.append(f"{ticker} ({region}) tried={[c[0] for c in cands]}")
            print(f"  [{i}/{len(items)}] {ticker:24s} -> 실패 {[c[0] for c in cands]}", flush=True)
        time.sleep(0.15)

    # 캐시 → prices.json 취합
    prices = {}
    for fn in os.listdir(CACHE):
        if fn.endswith(".json"):
            rec = json.load(open(os.path.join(CACHE, fn), encoding="utf-8"))
            prices[rec["ticker"]] = {"yf": rec["yf"], "ccy": rec["ccy"],
                                     "dates": rec["dates"], "close": rec["close"]}
    json.dump(prices, open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False)

    print(f"\n[완료] 성공/캐시 {ok} · 매핑불가 {skip} · 다운로드실패 {fail}")
    print(f"prices.json: {len(prices)}종목, {os.path.getsize(OUT_JSON)/1e6:.1f}MB")
    if skipped:
        print(f"\n매핑불가({len(skipped)}):")
        for x in skipped[:40]:
            print("  -", x)
        if len(skipped) > 40:
            print(f"  ... 외 {len(skipped)-40}")
    if failed:
        print(f"\n다운로드실패({len(failed)}):")
        for x in failed[:40]:
            print("  -", x)
        if len(failed) > 40:
            print(f"  ... 외 {len(failed)-40}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
