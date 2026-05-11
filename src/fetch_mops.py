"""
MOPS 公開資訊觀測站 — 直連 FinMind 沒同步時的 fallback.

兩步流程:
1. fetch_q1_announcements_for_date(date_ad)
   → 撈該日所有公司「第31款」公告中, 主旨含「第1季合併財務報告」者.
   返回 list[{co_id, name, TYPEK, seq_no, spoke_date, spoke_time}]
2. fetch_announcement_detail(item)
   → 撈單筆公告詳細 (step=2 神奇值), regex 抽 EPS + 完整損益表.

設計考量:
- mopsov.twse.com.tw 子網域比 mops.twse.com.tw 友善 (海外鏡像, bot 防護較鬆)
- 需先 GET 主頁取 cookie 後再 POST
- 個股 detail 用 step=2 (step=0/1/3 都不會回 EPS row)
- 民國年: ROC = AD - 1911

整合點: main.py 在 fetch_eps 完成後, 比對 MOPS 公告 vs FinMind snapshot,
找 FinMind 還沒 latest_quarter=2026Q1 但 MOPS 已有公告的, 把 EPS 注入 eps_data.
"""
import http.cookiejar
import re
import time
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Optional

MOPS_BASE = 'https://mopsov.twse.com.tw/mops/web'
USER_AGENT = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
              'AppleWebKit/537.36 (KHTML, like Gecko) '
              'Chrome/120.0.0.0 Safari/537.36')


def _new_session():
    """建立帶 cookie jar 的 urllib opener, 預先 GET 主頁拿 session cookie."""
    jar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
    opener.addheaders = [
        ('User-Agent', USER_AGENT),
        ('Accept', 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'),
        ('Accept-Language', 'zh-TW,zh;q=0.9'),
    ]
    try:
        opener.open(f'{MOPS_BASE}/t05st01', timeout=15)
    except Exception:
        pass  # 沒有 cookie 也試看看
    return opener


def _post(opener, form: dict, retries: int = 3, sleep: float = 1.0) -> Optional[str]:
    """POST 到 ajax_t05st01, 回應 HTML 字串. 失敗 retry."""
    url = f'{MOPS_BASE}/ajax_t05st01'
    data = urllib.parse.urlencode(form).encode()
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Referer': f'{MOPS_BASE}/t05st01',
        'X-Requested-With': 'XMLHttpRequest',
    }
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, data=data, headers=headers)
            with opener.open(req, timeout=30) as r:
                return r.read().decode('utf-8', errors='ignore')
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(sleep * (attempt + 1))
            else:
                print(f'  [MOPS] POST failed all {retries} retries: {e}')
                return None
    return None


def _ad_to_roc_date(date_ad: str) -> tuple[str, str, str]:
    """'2026-05-11' → ('115', '5', '11'). 月份不補零 (MOPS 接受 '5')."""
    y, m, d = date_ad.split('-')
    return str(int(y) - 1911), str(int(m)), str(int(d))


def fetch_q1_announcements_for_date(date_ad: str, quarter: int = 1) -> list[dict]:
    """掃指定日期所有第31款公告中, 主旨含「第N季合併財務報告」者.

    返回 list of {co_id, name, TYPEK, seq_no, spoke_date, spoke_time}.
    """
    opener = _new_session()
    year, month, day = _ad_to_roc_date(date_ad)
    form = {
        'encodeURIComponent': '1', 'step': '0', 'firstin': 'true', 'off': '1',
        'TYPEK': 'all', 'year': year, 'month': month,
        'b_date': day, 'e_date': day, 'subject1': '31',
    }
    html = _post(opener, form)
    if not html:
        return []
    # 找符合「第 N 季合併財務報告」主旨的 row, 從 onclick 抽出參數
    # 接受 "第1季" / "第一季" / "Q1" 多種寫法
    quarter_digit = ['', '1', '2', '3', '4'][quarter]
    quarter_chinese = ['', '一', '二', '三', '四'][quarter]
    quarter_re = re.compile(
        f'第\\s*[{quarter_digit}{quarter_chinese}]\\s*季\\s*[暨之]?[\\s\\S]{{0,15}}?合併財務報告'
    )
    results = []
    for row in re.findall(r'<tr[^>]*>(.+?)</tr>', html, re.S):
        if '合併財務報告' not in row:
            continue
        # 主旨在某個 <td> 裡, 把 row 內所有文字串起來檢查
        row_text = re.sub(r'<[^>]+>', ' ', row)
        row_text = re.sub(r'\s+', ' ', row_text)
        if not quarter_re.search(row_text):
            continue
        code_m = re.search(r"co_id\.value\s*=\s*['\"](\d+)['\"]", row)
        seq_m = re.search(r"seq_no\.value\s*=\s*['\"](\d+)['\"]", row)
        time_m = re.search(r"spoke_time\.value\s*=\s*['\"](\d+)['\"]", row)
        date_m = re.search(r"spoke_date\.value\s*=\s*['\"](\d+)['\"]", row)
        type_m = re.search(r"TYPEK\.value\s*=\s*['\"](\w+)['\"]", row)
        # 抓名稱 (第二個 td 是名稱)
        cells = re.findall(r"<td[^>]*>\s*(?:&nbsp;)?\s*([^<]*?)\s*</td>", row)
        name = cells[1].strip() if len(cells) > 1 else ''
        if code_m and seq_m and date_m and time_m and type_m:
            results.append({
                'co_id': code_m.group(1),
                'name': name,
                'TYPEK': type_m.group(1),
                'seq_no': seq_m.group(1),
                'spoke_date': date_m.group(1),
                'spoke_time': time_m.group(1),
                'opener': opener,  # reuse session
            })
    return results


def fetch_announcement_detail(item: dict, throttle_sec: float = 0.5) -> Optional[dict]:
    """撈單筆公告 step=2 詳細, regex 抽 EPS + 累計損益.

    回傳 {eps, revenue, gross_profit, operating_income, nonop,
          pretax_income, net_income, parent_income, quarter_end} 或 None.
    """
    opener = item.get('opener') or _new_session()
    year, _, _ = _ad_to_roc_date(_roc_to_ad(item['spoke_date']))
    form = {
        'step': '2', 'firstin': 'true', 'off': '1',
        'TYPEK': item['TYPEK'],
        'seq_no': item['seq_no'],
        'spoke_date': item['spoke_date'],
        'spoke_time': item['spoke_time'],
        'co_id': item['co_id'],
        'year': year, 'month': str(int(item['spoke_date'][3:5])),  # 從 ROC 日期取月份
        'b_date': str(int(item['spoke_date'][5:7])), 'e_date': str(int(item['spoke_date'][5:7])),
        'subject1': '31',
    }
    time.sleep(throttle_sec)
    html = _post(opener, form)
    if not html or '基本每股盈餘' not in html:
        return None
    text = re.sub(r'<[^>]+>', ' ', html)
    text = re.sub(r'\s+', ' ', text).replace('&nbsp;', ' ').strip()

    def _num(pat, cast=int):
        m = re.search(pat, text)
        if not m:
            return None
        v = m.group(1).strip().replace(',', '')  # 處理 "8,580,758"
        try:
            return cast(v)
        except (ValueError, TypeError):
            return None

    # 注意: MOPS 數字有時帶千分位逗號 (上市公司常見), 上櫃通常無
    NUM_RE = r'([\-\d,]+)'
    out = {
        'eps': _num(r'累計至本期止基本每股盈餘[^:：]*[(（][^)）]+[)）]\s*[(（]\s*元\s*[)）]\s*[:：]?\s*([\-\d.,]+)', float),
        'revenue': _num(rf'累計至本期止營業收入\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'gross_profit': _num(rf'累計至本期止營業毛利[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'operating_income': _num(rf'累計至本期止營業利益[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'pretax_income': _num(rf'累計至本期止稅前淨利[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'net_income': _num(rf'累計至本期止本期淨利[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'parent_income': _num(rf'累計至本期止歸屬於母公司業主淨利[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
    }
    # 推導 quarter_end 從財報期間
    m = re.search(r'起訖日期[^:：]*[:：]?\s*(\d{3}/\d{2}/\d{2})\s*~\s*(\d{3}/\d{2}/\d{2})', text)
    if m:
        out['quarter_end'] = _roc_to_ad(m.group(2).replace('/', ''))

    # 衍生指標
    rev = out.get('revenue')
    if rev:
        # FinMind 是仟元 但已存全額 (e.g. 10173410 千元 → keep as 10173410000)
        # 我們這邊統一存 仟元 (跟 FinMind 一致), 看上下游用法
        # FinMind 實際存 value 是 "千元" 數字 (10173410 表示一百億)
        # → 跟我們 fetch_eps 對齊, 保留原樣
        out['gm_pct'] = round((out.get('gross_profit') or 0) / rev, 4) if out.get('gross_profit') else None
        out['opm_pct'] = round((out.get('operating_income') or 0) / rev, 4) if out.get('operating_income') else None
    # 業外 = 稅前 - 營業利益
    if out.get('pretax_income') is not None and out.get('operating_income') is not None:
        out['nonop'] = out['pretax_income'] - out['operating_income']
    # 業外比
    ni = out.get('parent_income') or out.get('net_income')
    if ni and ni != 0 and out.get('nonop') is not None:
        out['nonop_pct'] = round(out['nonop'] / ni, 4)
    out['eps_source'] = 'mops'
    return out


def _roc_to_ad(roc_yyyymmdd: str) -> str:
    """'20260511' or '1150511' (民國 yyyy/mm/dd) → '2026-05-11'.

    我們從 spoke_date 拿到的格式是 '20260511' (西元) 但有時候是民國 7 位.
    """
    s = roc_yyyymmdd.replace('/', '').replace('-', '')
    if len(s) == 7:  # ROC 7-digit (1150511 → 民國 115/05/11)
        roc_y = int(s[:3])
        return f'{roc_y + 1911}-{s[3:5]}-{s[5:7]}'
    elif len(s) == 8:  # AD 8-digit
        return f'{s[:4]}-{s[4:6]}-{s[6:8]}'
    return roc_yyyymmdd


def fetch_mops_supplement(date_ad: str, missing_stock_ids: set,
                          quarter: int = 1, throttle_sec: float = 0.5,
                          progress: bool = True) -> dict:
    """整合便利函式: 掃指定日期 MOPS Q{quarter} 公告, 對 missing_stock_ids 中
    存在的, 回 detail dict.

    回傳 {stock_id: {date: {eps, revenue, ...}}} (對齊 fetch_eps 格式).
    """
    if progress:
        print(f'[MOPS] 掃 {date_ad} Q{quarter} 合併財務報告公告...')
    announcements = fetch_q1_announcements_for_date(date_ad, quarter=quarter)
    if progress:
        print(f'  共 {len(announcements)} 檔今日公告 Q{quarter}')
    target = [a for a in announcements if a['co_id'] in missing_stock_ids]
    if progress:
        print(f'  其中 {len(target)} 檔是我們缺的 (FinMind 沒同步)')
    out = {}
    for i, a in enumerate(target, 1):
        if progress and i % 10 == 0:
            print(f'  [{i}/{len(target)}] 抓 {a["co_id"]} {a["name"]}...', flush=True)
        detail = fetch_announcement_detail(a, throttle_sec=throttle_sec)
        if not detail or detail.get('eps') is None:
            continue
        q_end = detail.get('quarter_end')
        if not q_end:
            continue
        # 對齊 fetch_eps 格式: {date: {eps, revenue, ...}}
        fin = {k: v for k, v in detail.items() if k != 'quarter_end'}
        out[a['co_id']] = {q_end: fin}
    if progress:
        print(f'[MOPS] 抓回 {len(out)} 檔有效資料')
    return out


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    # 自測: 看今天有哪些 Q1 公告 + 抓 8096
    ann = fetch_q1_announcements_for_date('2026-05-11')
    print(f'5/11 共 {len(ann)} 檔')
    for a in ann[:5]:
        print(f'  {a["co_id"]} {a["name"]} {a["TYPEK"]} {a["spoke_date"]}/{a["spoke_time"]}')
    # 找 8096
    eight96 = next((a for a in ann if a['co_id'] == '8096'), None)
    if eight96:
        print(f'\n=== 8096 擎亞 detail ===')
        d = fetch_announcement_detail(eight96)
        for k, v in (d or {}).items():
            print(f'  {k}: {v}')
