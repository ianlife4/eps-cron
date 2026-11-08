"""
MOPS 公開資訊觀測站 — 直連 FinMind 沒同步時的 fallback.

兩步流程:
1. fetch_announcements_for_date(date_ad, quarter)
   → 撈該日所有公司「第31款」公告中, 主旨是「第N季(合併/個別)財務報告/報表」者 (Q4=年度財報).
   返回 list[{co_id, name, TYPEK, seq_no, spoke_date, spoke_time, subject}]
2. fetch_announcement_detail(item)
   → 撈單筆公告詳細 (step=2 神奇值), regex 抽 EPS + 完整損益表.
3. to_single_quarter()
   ⚠ MOPS 第31款數字是「1月1日累計至本期止」且單位「仟元」; FinMind 是「單季」且單位「元」。
   Q2/Q3/Q4 必須 累計 − 同年前幾季(FinMind) 才是單季, 金額 ×1000 才對得上
   (2026-11-08 修: 原本寫死 quarter=1, Q2/Q3/年報季整個補抓失效)。

設計考量:
- mopsov.twse.com.tw 子網域比 mops.twse.com.tw 友善 (海外鏡像, bot 防護較鬆)
- 需先 GET 主頁取 cookie 後再 POST
- 個股 detail 用 step=2 (step=0/1/3 都不會回 EPS row)
- 民國年: ROC = AD - 1911

整合點: main.py 在 fetch_eps 完成後, 比對 MOPS 公告 vs FinMind snapshot,
找 FinMind 還沒有最新一季 (freshest_q) 但 MOPS 已有公告的, 換算成單季後注入 eps_data.
"""
import http.cookiejar
import re
import time
import urllib.parse
import urllib.request
from collections import Counter
from datetime import datetime
from typing import Optional

MOPS_BASE = 'https://mopsov.twse.com.tw/mops/web'
# MOPS 限流頁是 HTTP 200:「Overrun - 查詢過於頻繁,請稍後再試!! Too many query requests…」
# 不攔會被當成「沒資料」靜默漏抓 (CB 系統 mops_client 實測容忍度 ~1.2 次/秒)
PAT_OVERRUN = re.compile(r'Overrun\s*-\s*查詢過於頻繁|Too many query requests')
# 金額欄位 (MOPS 仟元 → ×1000 對齊 FinMind 元)
MONEY_KEYS = ('revenue', 'gross_profit', 'operating_income',
              'pretax_income', 'net_income', 'parent_income')
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


def _post(opener, form: dict, retries: int = 4, sleep: float = 1.0) -> Optional[str]:
    """POST 到 ajax_t05st01, 回應 HTML 字串. 失敗 / 被限流 (Overrun) 都 retry."""
    url = f'{MOPS_BASE}/ajax_t05st01'
    data = urllib.parse.urlencode(form).encode()
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Referer': f'{MOPS_BASE}/t05st01',
        'X-Requested-With': 'XMLHttpRequest',
    }
    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, data=data, headers=headers)
            with opener.open(req, timeout=30) as r:
                html = r.read().decode('utf-8', errors='ignore')
            if not PAT_OVERRUN.search(html):
                return html
            last_err = 'Overrun (查詢過於頻繁)'
            wait = 10 * (attempt + 1)
            print(f'  [MOPS] 被限流 Overrun → 等 {wait}s 重試')
            time.sleep(wait)
        except Exception as e:
            last_err = e
            if attempt < retries - 1:
                time.sleep(sleep * (attempt + 1))
    print(f'  [MOPS] POST failed all {retries} retries: {last_err}')
    return None


def _ad_to_roc_date(date_ad: str) -> tuple[str, str, str]:
    """'2026-05-11' → ('115', '5', '11'). 月份不補零 (MOPS 接受 '5')."""
    y, m, d = date_ad.split('-')
    return str(int(y) - 1911), str(int(m)), str(int(d))


def _quarter_subject_re(quarter: int) -> re.Pattern:
    """第 N 季財報公告主旨 regex.

    2026-11-06 實測 Q3 主旨寫法 (共 ~240 則): 第三季/第3季/年度第三季/第三季度/前三季 ×
    合併財務報告/財務報告/合併財務報表/個別財務報告/XBRL合併財務報告 — 約 1/4 不含「合併財務報告」,
    舊版只認「合併財務報告」會漏。Q4 = 年報「115年度(合併/個體)財務報告」(排除 上半年度/半年度、年度第N季)。
    """
    tail = r'[\s\S]{0,15}?財務報[告表]'
    if quarter == 4:
        return re.compile(r'(?<![上半])年度(?:(?!季)[\s\S]){0,12}?財務報[告表]'
                          r'|第\s*[4四]\s*季' + tail)
    d, c = '1234'[quarter - 1], '一二三四'[quarter - 1]
    alts = [rf'第\s*[{d}{c}]\s*季']
    if quarter == 2:
        alts.append(r'半年度')
    elif quarter == 3:
        alts.append(r'前三季')
    return re.compile(rf'(?:{"|".join(alts)}){tail}')


def fetch_announcements_for_date(date_ad: str, quarter: int = 1) -> list[dict]:
    """掃指定日期所有第31款公告中, 主旨是「第N季財務報告」者 (quarter=4 → 年度財報).

    返回 list of {co_id, name, TYPEK, seq_no, spoke_date, spoke_time, subject},
    同 co_id 可能多筆 (已依優先序排好: 合併優先、較晚優先)。
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
    quarter_re = _quarter_subject_re(quarter)
    results = []
    for row in re.findall(r'<tr[^>]*>(.+?)</tr>', html, re.S):
        if '財務報' not in row:
            continue
        # 主旨在某個 <td> 裡, 把 row 內所有文字串起來檢查
        row_text = re.sub(r'<[^>]+>', ' ', row)
        row_text = re.sub(r'\s+', ' ', row_text.replace('&nbsp;', ' ')).strip()
        if not quarter_re.search(row_text):
            continue
        # 「代子公司 XX 公告…財務報告」: co_id 是母公司但數字是子公司的 → 不能注入
        # 「年度自結財務報告」(1~2 月) 是未查核自結數, 不是正式財報 → 自結速報那條線處理
        if '子公司' in row_text or '自結' in row_text:
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
                'subject': row_text,
                'opener': opener,  # reuse session
            })
    # 同公司可能多則 (合併+個別、原公告+「更正…XBRL」、董事會日期預告) — 全留, 依優先序排:
    # 合併 > 個別/個體, 同類較晚的優先; 由 fetch_mops_supplement 逐則試到有 EPS 為止
    results.sort(key=lambda r: ('合併' in r['subject'], r['spoke_date'] + r['spoke_time']), reverse=True)
    return results


# 舊名相容
fetch_q1_announcements_for_date = fetch_announcements_for_date


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
        # 會計括號負數 "(539,748)" / "(1.11)" (2340 台亞) — 舊版解析不到 → 虧損股整筆漏掉
        neg = v[:1] in '(（' and v[-1:] in ')）'
        v = v.strip('()（） ')
        try:
            return -cast(v) if neg else cast(v)
        except (ValueError, TypeError):
            return None

    # 注意: MOPS 數字有時帶千分位逗號 (上市公司常見), 上櫃通常無; 負數有 "-123" 也有 "(123)"
    NUM_RE = r'([(（]\s*[\d,]+\s*[)）]|-?[\d,]+)'
    out = {
        'eps': _num(r'累計至本期止基本每股盈餘[^:：]*[(（][^)）]+[)）]\s*[(（]\s*元\s*[)）]\s*[:：]?\s*'
                    r'([(（]\s*[\d.,]+\s*[)）]|-?[\d.,]+)', float),
        'revenue': _num(rf'累計至本期止營業收入\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'gross_profit': _num(rf'累計至本期止營業毛利[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'operating_income': _num(rf'累計至本期止營業利益[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'pretax_income': _num(rf'累計至本期止稅前淨利[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'net_income': _num(rf'累計至本期止本期淨利[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
        'parent_income': _num(rf'累計至本期止歸屬於母公司業主淨利[^:：]*[(（][^)）]+[)）]\s*[(（]\s*仟元\s*[)）]\s*[:：]?\s*{NUM_RE}'),
    }
    # 推導 quarter_end 從財報期間
    # 分隔符有「~」也有「-」(2483 百容 115/01/01-115/09/30)
    m = re.search(r'起訖日期[^:：]*[:：]?\s*(\d{3}/\d{2}/\d{2})\s*[~～\-－至]\s*(\d{3}/\d{2}/\d{2})', text)
    if m:
        out['period_start'] = _roc_to_ad(m.group(1).replace('/', ''))
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


def to_single_quarter(cum: dict, q_end: str, existing: Optional[dict]) -> tuple[Optional[dict], Optional[str]]:
    """MOPS「1/1 累計至本期止、仟元」→ FinMind 口徑「單季、元」.

    Q1 = 累計本身; Q2/Q3/Q4 = 累計 − 同年前幾季 (取自 existing = 該股 FinMind 季資料)。
    前幾季缺 EPS 就放棄 (回 None, 原因) — 寧可等 FinMind 隔天同步, 也不注入錯的單季數。
    EPS 相減會跟官方單季 EPS 差 0.0x (加權股數不同), FinMind 隔天同步後即以官方值為準。
    """
    if not q_end or q_end[5:7] not in ('03', '06', '09', '12'):
        return None, '無財報期間'
    if cum.get('eps') is None:
        return None, '明細無EPS/抓取失敗'
    y, m = int(q_end[:4]), int(q_end[5:7])
    q = m // 3
    fin = {k: (cum[k] * 1000 if cum.get(k) is not None else None) for k in MONEY_KEYS}
    fin['eps'] = cum.get('eps')
    # 期間起日不是 1/1 (極少數單季版) → 本身就是單季
    cumulative = not cum.get('period_start') or cum['period_start'].endswith('-01-01')
    if q > 1 and cumulative:
        existing = existing or {}
        prior_dates = [f'{y}-{md}' for md in ('03-31', '06-30', '09-30')[:q - 1]]
        prior = [existing.get(d) if isinstance(existing.get(d), dict) else {} for d in prior_dates]
        missing = [d for d, p in zip(prior_dates, prior) if p.get('eps') is None]
        if missing:
            return None, '缺前季EPS'
        fin['eps'] = round(fin['eps'] - sum(p['eps'] for p in prior), 2)
        for k in MONEY_KEYS:
            if fin[k] is None or any(p.get(k) is None for p in prior):
                fin[k] = None
            else:
                fin[k] = fin[k] - sum(p[k] for p in prior)
    # 衍生指標 (同 fetch_eps 口徑)
    rev = fin.get('revenue')
    gp, op = fin.get('gross_profit'), fin.get('operating_income')
    fin['gm_pct'] = round(gp / rev, 4) if rev and gp is not None else None
    fin['opm_pct'] = round(op / rev, 4) if rev and op is not None else None
    pt = fin.get('pretax_income')
    fin['nonop'] = pt - op if pt is not None and op is not None else None
    ni = fin.get('parent_income') or fin.get('net_income')
    fin['nonop_pct'] = round(fin['nonop'] / ni, 4) if ni and fin['nonop'] is not None else None
    fin['eps_source'] = 'mops_ytd' if (q > 1 and cumulative) else 'mops'
    return fin, None


def fetch_mops_supplement(date_ad: str, missing_stock_ids: set,
                          quarter: int = 1, throttle_sec: float = 0.8,
                          progress: bool = True,
                          existing_data: Optional[dict] = None,
                          expected_q_end: Optional[str] = None) -> dict:
    """整合便利函式: 掃指定日期 MOPS Q{quarter} 財報公告, 對 missing_stock_ids 中
    存在的, 抓 detail 並換算成單季 (見 to_single_quarter).

    existing_data: {stock_id: {date: fin}} — FinMind 現有季資料, Q2~Q4 扣前季用。
    expected_q_end: 只收這一季 (e.g. '2026-09-30'), 其他期別 (補發/更正舊季) 丟掉。
    回傳 {stock_id: {date: {eps, revenue, ...}}} (對齊 fetch_eps 格式, 金額單位元)。
    """
    label = '年度' if quarter == 4 else f'Q{quarter}'
    if progress:
        print(f'[MOPS] 掃 {date_ad} {label} 財務報告公告...')
    announcements = fetch_announcements_for_date(date_ad, quarter=quarter)
    groups: dict[str, list] = {}
    for a in announcements:
        groups.setdefault(a['co_id'], []).append(a)   # 已依優先序排好
    if progress:
        print(f'  共 {len(groups)} 檔今日公告 {label} ({len(announcements)} 則)')
    target = {sid: c for sid, c in groups.items() if sid in missing_stock_ids}
    if progress:
        print(f'  其中 {len(target)} 檔是我們缺的 (FinMind 沒同步)')
    out = {}
    skipped = Counter()
    for i, (sid, cands) in enumerate(target.items(), 1):
        if progress and i % 10 == 0:
            print(f'  [{i}/{len(target)}] 抓 {sid} {cands[0]["name"]}...', flush=True)
        why = None
        for a in cands:   # 逐則試到拿到本季 EPS 為止 (更正/預告公告沒有 EPS)
            detail = fetch_announcement_detail(a, throttle_sec=throttle_sec)
            if not detail or detail.get('eps') is None:
                why = '明細無EPS/抓取失敗'
                continue
            q_end = detail.get('quarter_end')
            if expected_q_end and q_end and q_end != expected_q_end:
                why = '期別不符'
                continue
            fin, why = to_single_quarter(detail, q_end, (existing_data or {}).get(sid))
            if fin is None:
                continue
            # 對齊 fetch_eps 格式: {date: {eps, revenue, ...}}
            out[sid] = {q_end: fin}
            break
        if sid not in out:
            skipped[why] += 1
    if progress:
        extra = f' (略過 {dict(skipped)})' if skipped else ''
        print(f'[MOPS] 抓回 {len(out)} 檔有效資料{extra}')
    return out


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    # 自測: 看 5/11 有哪些 Q1 公告 + 抓 8096
    ann = fetch_announcements_for_date('2026-05-11', quarter=1)
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
