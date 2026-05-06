"""
比較邏輯：
- YoY (本季 vs 去年同季)
- 累計達成率 (今年累計 vs 去年全年)
"""
from typing import Optional


def quarter_label(date_str: str) -> str:
    """'2025-03-31' -> '2025Q1'"""
    y, m, _ = date_str.split('-')
    q = (int(m) - 1) // 3 + 1
    return f'{y}Q{q}'


def same_quarter_last_year(date_str: str) -> str:
    """'2025-03-31' -> '2024-03-31'"""
    y, m, d = date_str.split('-')
    return f'{int(y) - 1}-{m}-{d}'


def find_latest_quarter(eps_data: dict) -> Optional[str]:
    """從 {date: {eps, ...}} 中找最新已公告的季。"""
    if not eps_data:
        return None
    valid_dates = [d for d, fin in eps_data.items() if fin.get('eps') is not None]
    return max(valid_dates) if valid_dates else None


def compute_yoy(latest_eps: float, prior_eps: Optional[float]) -> Optional[dict]:
    """YoY 比較。
    回傳 {delta, pct} 或 None。
    """
    if prior_eps is None or latest_eps is None:
        return None
    delta = round(latest_eps - prior_eps, 2)
    if abs(prior_eps) < 0.01:
        # 去年接近 0，不算百分比
        return {'delta': delta, 'pct': None, 'note': 'prior≈0'}
    pct = (latest_eps - prior_eps) / abs(prior_eps)
    return {'delta': delta, 'pct': round(pct, 4)}


def compute_accumulated(eps_data: dict, year: int) -> dict:
    """計算指定年份「截至最新已公告季」的累計 EPS。
    回傳 {value, quarters_count, latest_quarter}
    """
    same_year_dates = sorted([d for d in eps_data if d.startswith(str(year)) and eps_data[d].get('eps') is not None])
    if not same_year_dates:
        return {'value': None, 'quarters_count': 0, 'latest_quarter': None}
    total = sum(eps_data[d]['eps'] for d in same_year_dates)
    return {
        'value': round(total, 2),
        'quarters_count': len(same_year_dates),
        'latest_quarter': same_year_dates[-1],
    }


def compute_full_year(eps_data: dict, year: int) -> Optional[float]:
    """指定年份全年 EPS = 4 季加總（要 4 季都齊）"""
    yr_dates = sorted([d for d in eps_data if d.startswith(str(year)) and eps_data[d].get('eps') is not None])
    if len(yr_dates) < 4:
        return None
    return round(sum(eps_data[d]['eps'] for d in yr_dates), 2)


def analyze_one(stock_id: str, eps_data: dict, year_now: int = 2026) -> dict:
    """單檔分析。
    產出: {latest_quarter, latest_eps, yoy, accumulated, vs_prior_year_full, ...}
    """
    if not eps_data:
        return {'stock_id': stock_id, 'has_data': False}
    latest_date = find_latest_quarter(eps_data)
    if not latest_date:
        return {'stock_id': stock_id, 'has_data': False}

    latest_eps = eps_data[latest_date].get('eps')
    latest_q = quarter_label(latest_date)

    # YoY
    yoy_date = same_quarter_last_year(latest_date)
    yoy_eps = eps_data.get(yoy_date, {}).get('eps')
    yoy = compute_yoy(latest_eps, yoy_eps)

    # 累計達成率（今年累計 vs 去年全年）
    latest_year = int(latest_date[:4])
    accum = compute_accumulated(eps_data, latest_year)
    prior_full = compute_full_year(eps_data, latest_year - 1)
    achievement_pct = None
    if accum['value'] is not None and prior_full and prior_full != 0:
        achievement_pct = round(accum['value'] / prior_full, 4)
    elif accum['value'] is not None and prior_full == 0:
        achievement_pct = None
    elif accum['value'] is not None and prior_full is not None and prior_full < 0:
        # 去年虧損 → 不算 % 但標註
        achievement_pct = 'prior_loss'

    # QoQ (上一季)
    sorted_dates = sorted([d for d in eps_data if eps_data[d].get('eps') is not None])
    idx = sorted_dates.index(latest_date) if latest_date in sorted_dates else -1
    qoq = None
    if idx > 0:
        prev_eps = eps_data[sorted_dates[idx - 1]].get('eps')
        qoq = compute_yoy(latest_eps, prev_eps)

    fin = eps_data[latest_date]
    return {
        'stock_id': stock_id,
        'has_data': True,
        'latest_quarter': latest_q,
        'latest_date': latest_date,
        'latest_eps': latest_eps,
        'yoy_eps': yoy_eps,
        'yoy': yoy,
        'qoq': qoq,
        'accumulated': accum,
        'prior_year_full': prior_full,
        'achievement_pct': achievement_pct,
        'gm_pct': fin.get('gm_pct'),
        'opm_pct': fin.get('opm_pct'),
        'nonop_pct': fin.get('nonop_pct'),
    }


if __name__ == '__main__':
    # 自測
    sample = {
        '2024-03-31': {'eps': 1.04, 'gm_pct': 0.07, 'opm_pct': 0.03},
        '2024-06-30': {'eps': 0.14},
        '2024-09-30': {'eps': 0.39},
        '2024-12-31': {'eps': 1.83},
        '2025-03-31': {'eps': 1.27, 'gm_pct': 0.07, 'opm_pct': 0.04},
        '2025-06-30': {'eps': 1.28},
        '2025-09-30': {'eps': 0.65},
        '2025-12-31': {'eps': 5.58},
    }
    r = analyze_one('5386', sample)
    import json
    print(json.dumps(r, ensure_ascii=False, indent=2))
