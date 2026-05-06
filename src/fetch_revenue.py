"""
FinMind 月營收爬蟲
"""
import json
import time
import urllib.parse
import urllib.request
from typing import Optional

FINMIND_API = 'https://api.finmindtrade.com/api/v4/data'


def _http_get(url: str, params: dict, retries: int = 3, timeout: int = 30) -> Optional[dict]:
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
                return None


def fetch_monthly_revenue(stock_id: str, start_date: str, end_date: str) -> dict:
    """撈單檔月營收。
    回傳 {YYYY-MM: {revenue, year, month}} (用 revenue_year/revenue_month 作為實際月份)
    """
    data = _http_get(FINMIND_API, {
        'dataset': 'TaiwanStockMonthRevenue',
        'data_id': stock_id,
        'start_date': start_date,
        'end_date': end_date,
    })
    if not data or 'data' not in data:
        return {}
    result = {}
    for d in data['data']:
        y = d.get('revenue_year')
        m = d.get('revenue_month')
        if y is None or m is None:
            continue
        ym = f'{y}-{m:02d}'
        result[ym] = {
            'revenue': d.get('revenue'),
            'year': y,
            'month': m,
        }
    return result


def analyze_revenue(revenue_data: dict) -> dict:
    """分析月營收 — 自己算 YoY + 累計 (FinMind 只給單月原始營收)
    輸入: {YYYY-MM: {revenue, year, month}}
    回傳: {has_data, latest_ym, revenue, yoy_rev, mom_yoy_pct, accum_rev, accum_yoy, accum_yoy_pct}
    """
    if not revenue_data:
        return {'has_data': False}
    sorted_ym = sorted(revenue_data.keys())
    latest_ym = sorted_ym[-1]
    cur = revenue_data[latest_ym]
    cur_year = cur['year']
    cur_month = cur['month']
    cur_rev = cur['revenue']

    # 去年同期單月
    prior_ym = f'{cur_year - 1}-{cur_month:02d}'
    yoy_rev = revenue_data.get(prior_ym, {}).get('revenue')

    # 今年累計 (1月到當月)
    accum_rev = sum(revenue_data[f'{cur_year}-{m:02d}'].get('revenue') or 0
                    for m in range(1, cur_month + 1)
                    if f'{cur_year}-{m:02d}' in revenue_data)

    # 去年累計 (1月到當月)
    accum_yoy = sum(revenue_data[f'{cur_year - 1}-{m:02d}'].get('revenue') or 0
                    for m in range(1, cur_month + 1)
                    if f'{cur_year - 1}-{m:02d}' in revenue_data)

    mom_yoy_pct = None
    if cur_rev is not None and yoy_rev is not None and yoy_rev != 0:
        mom_yoy_pct = round((cur_rev - yoy_rev) / abs(yoy_rev), 4)

    accum_yoy_pct = None
    if accum_rev > 0 and accum_yoy != 0:
        accum_yoy_pct = round((accum_rev - accum_yoy) / abs(accum_yoy), 4)

    return {
        'has_data': True,
        'latest_ym': latest_ym,
        'revenue': cur_rev,
        'yoy_rev': yoy_rev,
        'mom_yoy_pct': mom_yoy_pct,
        'accum_rev': accum_rev if accum_rev > 0 else None,
        'accum_yoy': accum_yoy if accum_yoy > 0 else None,
        'accum_yoy_pct': accum_yoy_pct,
    }


def fetch_batch_monthly(stock_ids: list, start_date: str, end_date: str,
                       throttle_sec: float = 0.3, progress: bool = True) -> dict:
    results = {}
    total = len(stock_ids)
    for i, sid in enumerate(stock_ids):
        if progress and i % 50 == 0:
            print(f'  [revenue {i}/{total}] {sid}...', flush=True)
        results[sid] = fetch_monthly_revenue(sid, start_date, end_date)
        time.sleep(throttle_sec)
    return results


if __name__ == '__main__':
    # 自測
    print('Test: 撈 5386 青雲 月營收 2024-2025...')
    data = fetch_monthly_revenue('5386', '2024-01-01', '2025-12-31')
    for date, info in sorted(data.items()):
        rev = info.get('revenue')
        rev_str = f'{rev/1000:.0f}K' if rev else 'N/A'
        print(f'  {date}: {rev_str}  YoY={info.get("yoy")}  MoM={info.get("mom")}')
