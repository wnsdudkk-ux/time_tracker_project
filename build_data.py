"""
모든 ETF의 xlsx → 단일 data.js (window.ETF_DATA = {...}) 빌드.
표준 라이브러리 + openpyxl 만 사용.

────────────────────────────────────────────────────────────────────
data.js 스키마 (schema_version 2)

window.ETF_DATA = {
  generated_at: "...",
  schema_version: 2,
  inferred_holidays: ["YYYY-MM-DD", ...],   # 전 ETF 공통 누락 → 휴장 추정일
  etfs: [{
    idx, name,
    meta: {first_date, last_date, n_dates, n_tickers},
    dates: [...],                            # 데이터가 존재하는 날짜(오름차순)
    tickers: {ticker: {name, region}},
    series: {ticker: {
      weight: [...],   # 공시 비중(%) 원본 그대로 (소수 2자리 반올림값)
      qty:    [...],   # 수량
      value:  [...],   # 평가금액(원)
      act:    [[i, aw], ...],  # ★ 일별 '능동 비중 변화' 희소 목록
                               #   i  : dates 인덱스. dates[i-1]→dates[i] 구간의 변화
                               #   aw : 능동 비중 변화(%p). 가격·환율 효과를 제거한,
                               #        운용역의 매매로 인한 비중 변화만 분리한 값
                               #   |aw| < ACT_FLOOR 이고 신규편입/청산이 아니면 생략
    }},
    aum:  [...],   # 일별 순자산(원) = Σ평가금액
    ret:  [...],   # 일별 포트폴리오 추정 수익률(%) (첫날 null)
    flow: [...],   # 일별 추정 설정/환매 순유입(원) = AUM - 전일AUM×(1+ret) (첫날 null)
    splits: [[i, ticker, ratio], ...],  # 액면분할 추정 이벤트
    weight_sums: [...], missing_dates: [...], invalid_dates: [...],
  }, ...]
}

프런트엔드 사용법:
  · 기간 [a, b] 동안 종목 t의 능동 비중 변화 = Σ aw  (a < i ≤ b 인 act 항목)
  · 능동 매매 금액(원)   = aw/100 × aum[i]            (일별), 기간은 합산
  · 신규 편입/전량 청산  = weight 배열에서 null↔값 전환으로 판별 (act에 항상 포함됨)
  · 가격 효과(수동 변화) = (비중 변화 총량) − (능동 변화 합)

능동 비중 변화의 정의:
  전일 포트폴리오를 그대로 들고 거래하지 않았을 때의 오늘 비중(w_passive)과
  실제 오늘 비중의 차이. w_passive = 전일비중 × 종목수익률 ÷ 포트수익률.
  종목수익률은 내재가격(평가금액÷수량)으로 산출하므로 환율 효과가 포함되며,
  비중은 규모 불변이므로 설정/환매(자금 유출입)에는 영향받지 않는다.
  단, 현금형 설정 직후 미투입 현금이 일시적으로 '전 종목 능동 축소 + 현금 능동
  확대'로 잡힐 수 있다(수일 내 재투입되면 기간 합산에서 상쇄).
────────────────────────────────────────────────────────────────────
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

HERE      = os.path.dirname(os.path.abspath(__file__))
ROOT      = os.path.join(HERE, "data")
OUT_JS    = os.path.join(HERE, "data.js")

DATE_RE   = re.compile(r"(\d{4}-\d{2}-\d{2})")
CACHE_VER = 2          # 파싱/정규화 로직이 바뀌면 +1 → 캐시 전체 무효화
ACT_FLOOR = 0.003      # |능동 비중 변화| 저장 하한(%p). 신규편입/청산은 무조건 저장
SPLIT_NUMS = [1.5, 2, 2.5, 3, 4, 5, 6, 7, 8, 10, 15, 20, 25, 50]
SPLIT_RATIOS = sorted(SPLIT_NUMS + [1.0 / x for x in SPLIT_NUMS])

# Bloomberg 식 'XX EQUITY' 접미사 → 지역
EQUITY_SUFFIX_REGION = {
    "US": "US", "HK": "HK", "JP": "JP", "JT": "JP",
    "CH": "CN", "C1": "CN", "C2": "CN", "CG": "CN",
    "KS": "KR", "KQ": "KR", "TT": "TW",
}
EQUITY_SUFFIX_RE = re.compile(r"\s([A-Z0-9]{2})\s+EQUITY$")


# ──────────────────────────────────────────────
# 1. 셀 값 파싱·정규화
# ──────────────────────────────────────────────

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


def compact_num(v, ndigits=None):
    """JSON 용량 절감: 반올림 후 정수면 int 로."""
    if v is None:
        return None
    if ndigits is not None:
        v = round(v, ndigits)
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def normalize_ticker(code, name):
    name_s = str(name).strip() if name is not None else ""
    if code is None or (isinstance(code, str) and not code.strip()):
        if name_s == "현금":
            return "__CASH__"
        # 코드 없는 비현금 행: 'None' 문자열로 뭉개지 말고 종목명 기반 키 부여
        return f"NAME::{name_s}"
    if isinstance(code, (int, float)):
        return f"{int(code):06d}"
    return str(code).strip()


def classify_region(ticker, name):
    if ticker == "__CASH__":
        return "Cash"
    m = EQUITY_SUFFIX_RE.search(ticker)
    if m:
        return EQUITY_SUFFIX_REGION.get(m.group(1), "Other")
    if re.fullmatch(r"\d{6}", ticker):          # 일반 한국 종목/ETF 코드
        return "KR"
    if re.fullmatch(r"\d{5}K", ticker):         # 한국 신형우선주 (예: 00104K)
        return "KR"
    if re.fullmatch(r"[0-9A-Z]{6}", ticker):
        return "Fund"
    return "Other"


# ──────────────────────────────────────────────
# 2. xlsx 읽기 (+ 파일 단위 캐시)
# ──────────────────────────────────────────────

HEADER_HINTS = ("종목코드", "종목명")


def read_xlsx(path):
    """헤더 행을 탐지해 그 아래를 파싱. 동일 티커 중복 행은 합산."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    merged = {}     # ticker → row dict (입력 순서 유지)
    header_seen = False
    for ri, row in enumerate(ws.iter_rows(values_only=True)):
        if row is None or len(row) < 5:
            continue
        if not header_seen:
            c0, c1 = str(row[0] or ""), str(row[1] or "")
            if HEADER_HINTS[0] in c0 and HEADER_HINTS[1] in c1:
                header_seen = True
                continue
            if ri >= 4:          # 5행 안에 헤더가 없으면 1행 헤더로 간주하고 진행
                header_seen = True
            else:
                continue
        code, name, qty, value, weight = row[0], row[1], row[2], row[3], row[4]
        if name is None:
            continue
        t = normalize_ticker(code, name)
        q, v, w = parse_number(qty), parse_number(value), parse_number(weight)
        if t in merged:
            m = merged[t]
            m["qty"]    = (m["qty"] or 0) + (q or 0) if (m["qty"] is not None or q is not None) else None
            m["value"]  = (m["value"] or 0) + (v or 0) if (m["value"] is not None or v is not None) else None
            m["weight"] = (m["weight"] or 0) + (w or 0) if (m["weight"] is not None or w is not None) else None
        else:
            merged[t] = {"ticker": t, "name": str(name).strip(),
                         "qty": q, "value": v, "weight": w}
    wb.close()
    return list(merged.values())


def load_holdings_cached(xlsx_dir, files):
    """파일 (mtime, size) 기준 캐시. 변경된 파일만 openpyxl 로 재파싱."""
    cache_path = os.path.join(xlsx_dir, ".parse_cache.json")
    cache = {}
    try:
        with open(cache_path, encoding="utf-8") as f:
            raw = json.load(f)
        if raw.get("ver") == CACHE_VER:
            cache = raw.get("files", {})
    except (OSError, ValueError):
        cache = {}

    out, new_cache, n_parsed = {}, {}, 0
    for f in files:
        base = os.path.basename(f)
        st = os.stat(f)
        key = [int(st.st_mtime), st.st_size]
        ent = cache.get(base)
        if ent and ent.get("k") == key:
            holdings = [
                {"ticker": h[0], "name": h[1], "qty": h[2], "value": h[3], "weight": h[4]}
                for h in ent["h"]
            ]
        else:
            holdings = read_xlsx(f)
            n_parsed += 1
        out[base] = holdings
        new_cache[base] = {
            "k": key,
            "h": [[h["ticker"], h["name"], h["qty"], h["value"], h["weight"]]
                  for h in holdings],
        }
    try:
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump({"ver": CACHE_VER, "files": new_cache}, f,
                      ensure_ascii=False, separators=(",", ":"))
    except OSError as e:
        print(f"    ! 캐시 저장 실패(무시): {e}")
    return out, n_parsed


# ──────────────────────────────────────────────
# 3. 일별 능동 분해
# ──────────────────────────────────────────────

def _detect_split(qty_ratio, price_ratio):
    """액면분할(병합) 추정. 반환: 분할배수 s 또는 None.

    판정(자금유출입 배율 추정이 필요 없는 가격 기반 방식):
      ① 내재가격이 하루 등락으로 보기 어려운 폭으로 점프했고(±40% 초과),
      ② 평가금액(수량비×가격비)은 연속이면(±50% 이내) → 분할.
    분할배수는 가격비의 역수에서 추정한다. 가격은 당일 매매가 섞여도
    오염되지 않으므로, 분할과 매매가 같은 날 겹쳐도 배수 추정이 유지된다.
    한계: 하루 ±40% 초과 급등락과 그에 상응하는 대량 매매가 '동시에' 일어난
    극단 사례는 분할로 오인할 수 있다(실무상 드묾)."""
    if qty_ratio is None or price_ratio is None or price_ratio <= 0:
        return None
    if 1 / 1.4 < price_ratio < 1.4:
        return None                      # 가격 점프 없음 → 분할 아님
    value_ratio = qty_ratio * price_ratio
    if not (1 / 1.5 < value_ratio < 1.5):
        return None                      # 평가금액 불연속 → 실제 급등락/대량 매매
    x = 1.0 / price_ratio
    best = min(SPLIT_RATIOS, key=lambda s: abs(x / s - 1))
    return best if abs(x / best - 1) < 0.10 else x   # 표에 없으면 관측비율 그대로


def decompose_active(dates, by_date):
    """인접한 두 데이터일 사이의 '능동 비중 변화'(%p)를 종목별로 산출.

    w_passive = 전일비중 × R_i / R_p  (무거래 가정 비중)
    aw        = 실제 오늘비중 − w_passive
    비중은 공시값이 아니라 평가금액/Σ평가금액으로 재계산해 사용(정밀도).
    R_i 는 내재가격(평가금액÷수량) 비율. 산출 불가 종목·전량청산 종목은 R_p 로 대체.
    """
    n = len(dates)
    aum   = [None] * n
    ret   = [None] * n
    flow  = [None] * n
    act   = {}                 # ticker → [[i, aw], ...]
    splits = []

    # 소스 평가금액 1일 글리치 보정(median-3): 분할일 수량만 조정되고 평가금액은 미조정되어
    # 하루만 ~20배로 튀었다 복귀하는 등의 오류를, 양 이웃보다 3배↑/0.34배↓로 튄 값을 세 값의
    # 중앙값으로 치환해 정상화한다. 현금·지수선물(주식수 기반 매매 대상 아님)은 제외.
    def _is_base(t):
        return t == "__CASH__" or "CASH" in t.upper() or re.search(r"\sINDEX$", t, re.I)

    val_series = {}            # ticker → [value per date or None]
    for i, ds in enumerate(dates):
        for h in by_date[ds]:
            val_series.setdefault(h["ticker"], [None] * n)[i] = h["value"]
    for t, arr in val_series.items():
        if _is_base(t):
            continue
        for i in range(1, n - 1):
            a, b, d = arr[i - 1], arr[i], arr[i + 1]
            if a and b and d and a > 0 and d > 0:
                r1, r2 = b / a, b / d
                if (r1 > 3 and r2 > 3) or (r1 < 0.34 and r2 < 0.34):
                    arr[i] = sorted([a, b, d])[1]

    snaps = []                 # 날짜별 {ticker: (qty, value)} — value 는 글리치 보정본
    for i, ds in enumerate(dates):
        snaps.append({h["ticker"]: (h["qty"], val_series[h["ticker"]][i]) for h in by_date[ds]})

    for i, ds in enumerate(dates):
        s = sum(v for _, v in snaps[i].values() if v is not None and v > 0)
        aum[i] = s if s > 0 else None

    for i in range(1, n):
        s0, s1 = snaps[i - 1], snaps[i]
        nav0, nav1 = aum[i - 1], aum[i]
        if not nav0 or not nav1:
            continue

        union = list(dict.fromkeys(list(s0.keys()) + list(s1.keys())))
        w0 = {t: (s0[t][1] / nav0 if t in s0 and s0[t][1] else 0.0) for t in union}
        w1 = {t: (s1[t][1] / nav1 if t in s1 and s1[t][1] else 0.0) for t in union}

        # 종목 수익률 R_i
        R = {}
        for t in union:
            if t == "__CASH__":
                R[t] = 1.0
                continue
            if t not in s0 or t not in s1:
                continue                       # 신규/청산 → R 불필요/미상
            q0, v0 = s0[t]
            q1, v1 = s1[t]
            if not (q0 and q1 and v0 and v1 and q0 > 0 and q1 > 0 and v0 > 0 and v1 > 0):
                continue
            price_ratio = (v1 / q1) / (v0 / q0)
            sp = _detect_split(q1 / q0, price_ratio)
            if sp is not None:
                R[t] = price_ratio * sp
                splits.append([i, t, compact_num(sp, 4)])
            else:
                R[t] = price_ratio

        known_w = sum(w0[t] for t in R if w0[t] > 0)
        if known_w <= 0:
            continue
        r_p = sum(w0[t] * R[t] for t in R if w0[t] > 0) / known_w

        ret[i]  = compact_num((r_p - 1) * 100, 3)
        flow[i] = compact_num(nav1 - nav0 * r_p, 0)

        for t in union:
            r_i = R.get(t, r_p)               # 미상·청산 종목은 시장 평균 수익 가정
            w_pass = w0[t] * r_i / r_p
            aw = (w1[t] - w_pass) * 100.0
            is_event = (w0[t] == 0) != (w1[t] == 0)   # 신규 편입 또는 전량 청산
            aw_r = compact_num(aw, 3)
            if is_event or abs(aw) >= ACT_FLOOR:
                if aw_r != 0 or is_event:
                    act.setdefault(t, []).append([i, aw_r])

        # 자체 점검: Σw_passive = 1 이어야 함(분해의 보존 법칙)
        chk = sum(w0[t] * R.get(t, r_p) / r_p for t in union)
        if abs(chk - 1) > 1e-9:
            print(f"    ! 분해 점검 실패 {dates[i]}: Σw_pass={chk:.12f}")

    return aum, ret, flow, act, splits


# ──────────────────────────────────────────────
# 4. ETF 단위 빌드
# ──────────────────────────────────────────────

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

    holdings_by_file, n_parsed = load_holdings_cached(xlsx_dir, files)
    print(f"[*] {etf['name']}: 파일 {len(files)}개 (신규 파싱 {n_parsed}, 캐시 {len(files) - n_parsed})")

    by_date = {}
    for f in files:
        base = os.path.basename(f)
        m = DATE_RE.search(base)
        if not m:
            continue
        h = holdings_by_file.get(base)
        if h:
            by_date[m.group(1)] = h

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
            series[t]["weight"][di] = compact_num(h["weight"], 4)
            series[t]["qty"][di]    = compact_num(h["qty"], 4)
            series[t]["value"][di]  = compact_num(h["value"], 0)
            if h["weight"] is not None:
                s += h["weight"]
        weight_sums.append(round(s, 2))

    aum, ret, flow, act, splits = decompose_active(dates, by_date)
    for t, pairs in act.items():
        series[t]["act"] = pairs

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
        "aum": aum,
        "ret": ret,
        "flow": flow,
        "splits": splits,
        "weight_sums": weight_sums,
        "missing_dates": [],      # finalize_missing() 에서 채움
        "invalid_dates": invalid,
    }


def finalize_missing(etfs):
    """주말 제외 누락일 중 '모든 ETF 가 동시에 비어 있는 날'은 휴장으로 추정해
    공통 목록(inferred_holidays)으로 빼고, 나머지만 ETF 별 missing_dates 로 둔다."""
    have = {e["idx"]: set(e["dates"]) for e in etfs}
    union_dates = set().union(*have.values()) if have else set()
    ranges = {e["idx"]: (Date.fromisoformat(e["meta"]["first_date"]),
                         Date.fromisoformat(e["meta"]["last_date"])) for e in etfs}

    holidays = set()
    for e in etfs:
        d0, d1 = ranges[e["idx"]]
        cur, missing = d0, []
        while cur <= d1:
            if cur.weekday() < 5:
                ds = cur.isoformat()
                if ds not in have[e["idx"]]:
                    covering = [i for i, (a, b) in ranges.items() if a <= cur <= b]
                    if ds not in union_dates and len(covering) >= 2:
                        holidays.add(ds)          # 전 ETF 공통 누락 → 휴장 추정
                    else:
                        missing.append(ds)        # 이 ETF 만 빠짐 → 진짜 누락
            cur += timedelta(days=1)
        e["missing_dates"] = missing
    return sorted(holidays)


# ──────────────────────────────────────────────
# 5. main
# ──────────────────────────────────────────────

def main() -> int:
    out = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "schema_version": 2,
        "inferred_holidays": [],
        "etfs": [],
    }
    for etf in ETFS:
        d = build_etf(etf)
        if d:
            out["etfs"].append(d)
    if not out["etfs"]:
        print("[!] 빌드된 ETF 없음. 먼저 다운로드를 실행하세요.")
        return 1

    out["inferred_holidays"] = finalize_missing(out["etfs"])

    with open(OUT_JS, "w", encoding="utf-8") as f:
        f.write("window.ETF_DATA = ")
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")

    size_mb = os.path.getsize(OUT_JS) / (1024 * 1024)
    print()
    print(f"[OK] {OUT_JS}  ({size_mb:.2f} MB)")
    print(f"     휴장 추정 제외일: {len(out['inferred_holidays'])}일")
    for e in out["etfs"]:
        m = e["meta"]
        n_act = sum(len(s.get("act", [])) for s in e["series"].values())
        print(f"   - {e['name']}: {m['n_dates']}일, {m['n_tickers']}종목, "
              f"능동변화 {n_act}건, 분할추정 {len(e['splits'])}건 "
              f"(누락 {len(e['missing_dates'])}일, 비중합 이상 {len(e['invalid_dates'])}일)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
