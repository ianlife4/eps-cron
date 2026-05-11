"""
FinMind EPS 爬蟲
- 撈台股財報 (EPS, GM, OPM, 業外) 季資料
- 支援上市/上櫃/創新板，興櫃/公發 fallback to MOPS
"""
import json
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Optional

FINMIND_API = 'https://api.finmindtrade.com/api/v4/data'
DATA_DIR = Path(__file__).parent.parent / 'data'

import os as _os
FINMIND_TOKEN = _os.environ.get('FINMIND_TOKEN', '').strip()

# FinMind 重要欄位對照
EPS_FIELDS = {
    'EPS': 'eps',
    'Revenue': 'revenue',
    'GrossProfit': 'gross_profit',
    'OperatingIncome': 'operating_income',
    'TotalNonoperatingIncomeAndExpense': 'nonop',
    'IncomeAfterTaxes': 'net_income',
    # 歸屬母公司淨利 — 用於 EPS fallback 計算 (FinMind 偶有 EPS row 同步落後)
    'EquityAttributableToOwnersOfParent': 'parent_income',
    'PreTaxIncome': 'pretax_income',
}

# 流通股數 cache (per-run, 避免重複 call TaiwanStockShareholding)
_SHARES_CACHE: dict[str, int] = {}


def _http_get(url: str, params: dict, retries: int = 3, timeout: int = 30) -> Optional[dict]:
    """HTTP GET with retry. 自動加 FinMind token."""
    if FINMIND_TOKEN and 'token' not in params:
        params = {**params, 'token': FINMIND_TOKEN}
    qs = urllib.parse.urlencode(params)
    full = f'{url}?{qs}'
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(full, timeout=timeout) as r:
                return json.loads(r.read())
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
            else:
                print(f'  ❌ {url} failed: {e}')
                return None
    return None


def fetch_stock_list(market: str = 'TWSE') -> list:
    """撈某市場全部股票清單。
    market: TWSE (上市) / TPEx (上櫃) / 等
    """
    data = _http_get(FINMIND_API, {
        'dataset': 'TaiwanStockInfo',
    })
    if not data or 'data' not in data:
        return []
    # 過濾市場
    return [s for s in data['data'] if s.get('type') == market or market == 'ALL']


def fetch_shares_outstanding(stock_id: str) -> Optional[int]:
    """撈 TaiwanStockShareholding 取最新 NumberOfSharesIssued (流通在外股數).

    用於 EPS fallback 計算: 當 FinMind EPS row 同步落後 (公告當天常見),
    用 歸母淨利 / 流通股數 自己算 EPS.

    結果 per-run cache (_SHARES_CACHE) 避免重複 API call.
    回傳 None 表示無資料.
    """
    if stock_id in _SHARES_CACHE:
        return _SHARES_CACHE[stock_id]
    # 取最近 30 天的最新一筆 (TaiwanStockShareholding 是日資料)
    from datetime import datetime, timedelta
    end = datetime.now().strftime('%Y-%m-%d')
    start = (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d')
    data = _http_get(FINMIND_API, {
        'dataset': 'TaiwanStockShareholding',
        'data_id': stock_id,
        'start_date': start,
        'end_date': end,
    })
    shares = None
    if data and data.get('data'):
        latest = max(data['data'], key=lambda x: x.get('date', ''))
        shares = latest.get('NumberOfSharesIssued')
    _SHARES_CACHE[stock_id] = shares
    return shares


def fetch_quarterly_eps(stock_id: str, start_date: str, end_date: str) -> dict:
    """撈單檔股票季 EPS 資料。
    回傳 {date: {eps, revenue, gross_profit, ...}}

    當某季 FinMind 缺 EPS row 但有 EquityAttributableToOwnersOfParent (歸母淨利):
    自動撈流通股數計算 EPS (fin['eps_source'] = 'computed', fin['eps_shares_used'] = N).
    處理 FinMind 公告當天 EPS 同步落後的常見問題.
    """
    data = _http_get(FINMIND_API, {
        'dataset': 'TaiwanStockFinancialStatements',
        'data_id': stock_id,
        'start_date': start_date,
        'end_date': end_date,
    })
    if not data or 'data' not in data:
        return {}
    result = {}
    for d in data['data']:
        date = d['date']
        ftype = d['type']
        if ftype not in EPS_FIELDS:
            continue
        if date not in result:
            result[date] = {}
        result[date][EPS_FIELDS[ftype]] = d['value']

    # === EPS fallback: 缺 EPS 但有歸母淨利 → 自算 ===
    needs_fallback = []
    for date, fin in result.items():
        if fin.get('eps') is None and (
            fin.get('parent_income') is not None or fin.get('net_income') is not None
        ):
            needs_fallback.append(date)
    if needs_fallback:
        shares = fetch_shares_outstanding(stock_id)
        if shares and shares > 0:
            for date in needs_fallback:
                fin = result[date]
                # 優先用歸母淨利 (台股 EPS 標準算法), 沒有再退到合併稅後
                income = fin.get('parent_income') or fin.get('net_income')
                if income is not None:
                    fin['eps'] = round(income / shares, 2)
                    fin['eps_source'] = 'computed'
                    fin['eps_shares_used'] = shares

    # 計算 GM%, OPM% 衍生指標
    for date, fin in result.items():
        rev = fin.get('revenue')
        if rev and rev != 0:
            fin['gm_pct'] = round(fin.get('gross_profit', 0) / rev, 4) if 'gross_profit' in fin else None
            fin['opm_pct'] = round(fin.get('operating_income', 0) / rev, 4) if 'operating_income' in fin else None
        # 業外比例 — 優先用歸母, 退到合併
        ni = fin.get('parent_income') or fin.get('net_income')
        nonop = fin.get('nonop')
        if ni and ni != 0 and nonop is not None:
            fin['nonop_pct'] = round(nonop / ni, 4)
    return result


def fetch_batch(stock_ids: list, start_date: str, end_date: str,
                throttle_sec: float = 0.3, progress: bool = True) -> dict:
    """批次撈取，回傳 {stock_id: {date: {eps, ...}}}."""
    results = {}
    total = len(stock_ids)
    for i, sid in enumerate(stock_ids):
        if progress and i % 50 == 0:
            print(f'  [{i}/{total}] {sid}...', flush=True)
        results[sid] = fetch_quarterly_eps(sid, start_date, end_date)
        time.sleep(throttle_sec)
    return results


def save_snapshot(data: dict, snapshot_name: str):
    """存 JSON snapshot 到 data/snapshots/"""
    out = DATA_DIR / 'snapshots' / f'{snapshot_name}.json'
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    return out


def load_snapshot(snapshot_name: str) -> Optional[dict]:
    """載入指定 snapshot。"""
    f = DATA_DIR / 'snapshots' / f'{snapshot_name}.json'
    if not f.exists():
        return None
    return json.loads(f.read_text(encoding='utf-8'))


if __name__ == '__main__':
    # 自測：撈 5386 青雲 2024-2025 兩年資料
    print('Test: 撈 5386 青雲 2024-2025...')
    data = fetch_quarterly_eps('5386', '2024-01-01', '2025-12-31')
    for date, fin in sorted(data.items()):
        print(f'  {date}: EPS={fin.get("eps")} GM%={fin.get("gm_pct")} OPM%={fin.get("opm_pct")}')
