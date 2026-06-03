"""
MOPS 自結速報 — 每天抓「自結 / 注意股 / 處置股」財務揭露裡的自結數 (領先官方財報).

來源: MOPS 重大訊息主旨全文檢索 (ajax_t51sb10, 第51款), 三組關鍵字去重:
  - 自結         → 公司自願公告月/季自結損益、營收、盈餘 (各家格式不同 = 模板B)
  - 注意交易資訊  → 達注意標準的財務業務揭露 (證交所統一表 = 模板A)
  - 處置交易資訊  → 達處置標準 (同模板A)
上市(sii)+上櫃(otc)各跑. 主旨不含「自結」的注意/處置股也會被抓到 (它們的自結數在內文).

明細解析 (ajax_t05st01 step=2):
  模板A (注意/處置股統一表): 純 regex. 抽「最近一月 = IFRS合併自結數」欄的 營收/稅前/稅後/EPS + 增減%.
                            ⚠ 同列也有「最近一季 = 查核數」(= 官方季報), 不要抓錯欄.
  模板B (自願自結): regex 打常見會計科目 (單位 仟元/百萬 normalize) + Claude Haiku 兜異常 prose.

金額一律 normalize 到「仟元」(對齊 FinMind/fetch_eps), EPS 單位「元」.
缺 EPS 時用 自結淨利(仟元)×1000 ÷ 流通股數 自算 (複用 fetch_eps.fetch_shares_outstanding).

輸出: list[dict], 每筆:
  {stock_id, name, announce_date, announce_time, subject, source_type,
   period_month, eps, eps_yoy, revenue, net_income, pretax, operating_income,
   eps_source, parse_method, ...}
"""
import re
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent))
from fetch_mops import _new_session, MOPS_BASE  # noqa: E402

# 搜尋關鍵字 → 預設 source_type (主旨還會再細分)
KEYWORDS = ['自結', '注意交易資訊', '處置交易資訊']
# 市場別由 t51sb10 的 KIND 參數控制 (TYPEK 被忽略):
#   L=上市(sii) / O=上櫃(otc) / R=興櫃(rotc). 實測補正: 原本兩組都用 L → 漏抓上櫃.
MARKET_KINDS = ['L', 'O', 'R']

# === 模板 A (注意/處置股統一表) 科目對照 (長名優先) ===
A_FIELDS = [
    ('revenue', ['營業收入淨額', '營業收入']),
    ('operating_income', ['營業利益', '營業損益']),
    ('pretax', ['稅前淨利', '稅前損益', '稅前純益']),
    # 很多注意股只報歸屬母公司淨利 (無「本期淨利」列), 一併納入
    ('net_income', ['本期淨利', '稅後淨利', '稅後純益', '本期損益',
                    '歸屬於母公司業主之淨利', '歸屬母公司業主淨利', '歸屬母公司', '歸屬母公']),
    ('eps', ['每股盈餘', '基本每股盈餘']),
]
# === 模板 B (自願自結) 科目對照 (長名優先, 避免子字串誤命中) ===
B_FIELDS = [
    ('revenue', ['合併營業收入淨額', '營業收入淨額', '合併營業收入', '營業收入']),
    ('operating_income', ['合併營業利益', '合併營業損益', '營業利益', '營業損益']),
    ('pretax', ['合併稅前損益', '合併稅前淨利', '稅前淨利', '稅前損益', '稅前純益']),
    ('net_income', ['合併稅後損益', '本期稅後淨利', '本期淨利', '稅後淨利', '稅後純益', '本期損益']),
    ('parent_income', ['歸屬於母公司業主之淨利', '歸屬於母公司', '母公司業主']),
    ('eps', ['基本每股盈餘', '每股盈餘', '每股稅後盈餘']),
]

# 數字 token: 括號負數 (25) / 一般 -25 / 帶 % / 缺值符號
NUM_RE = r'\([\d,]+\.?\d*%?\)|-?[\d,]+\.?\d*%?|--|—|N/A'


# ---------------------------------------------------------------- 數字工具
def _to_float(tok: Optional[str]) -> Optional[float]:
    """支援會計負數括號 '(25)' → -25, 千分位逗號, 結尾 %."""
    if not tok:
        return None
    t = tok.strip()
    neg = False
    if t.startswith('(') and t.endswith(')'):
        neg, t = True, t[1:-1]
    t = t.rstrip('%').replace(',', '').strip()
    if t in ('', '--', '—', 'N/A'):
        return None
    try:
        v = float(t)
        return -v if neg else v
    except ValueError:
        return None


def _pct(tok: Optional[str]) -> Optional[float]:
    """'141%' → 1.41 (比例); '(17,233%)' → -172.33."""
    v = _to_float(tok)
    return round(v / 100, 4) if v is not None else None


def _find_unit(text: str) -> Optional[str]:
    """抓單位字串. 處理「單位:新台幣佰萬」「(新台幣百萬元)」「(百萬)」「(元)」等。"""
    m = re.search(r'單位\s*[：:]\s*([^\s,，。;0-9]{1,8})', text)
    if m:
        return m.group(1)
    for mm in re.finditer(r'[（(]\s*(新?臺?台?幣?[仟千百佰萬億]*元?)\s*[)）]', text):
        u = mm.group(1)
        if any(c in u for c in '仟千百佰萬億') or u.endswith('元'):
            return u
    return None


def _normalize_amount(val: Optional[float], unit: Optional[str]) -> Optional[float]:
    """金額一律轉「仟元」. 百萬/佰萬→×1000, 億→×100000, 仟/千元→×1, 純元→/1000."""
    if val is None:
        return None
    u = (unit or '').replace('佰', '百')
    if '百萬' in u:
        return round(val * 1000)
    if '億' in u:
        return round(val * 100000)
    if '仟' in u or '千' in u:
        return round(val)
    if u.endswith('元'):  # 純「元」(極少見於金額)
        return round(val / 1000)
    return round(val)  # 未知 → 視為仟元


# ---------------------------------------------------------------- 搜尋層
def _search_once(opener, keyword: str, roc_year: int, month: int,
                 b_day: int, e_day: int, typek: str, kind: str) -> Optional[str]:
    kw = urllib.parse.quote(keyword)
    url = (f'{MOPS_BASE}/ajax_t51sb10?encodeURIComponent=1&firstin=true'
           f'&id=&key=&TYPEK={typek}&Stp=4&go=false&COMPANY_ID=&r1=1&KIND={kind}&CODE='
           f'&keyWord={kw}&year={roc_year}&month1={month}'
           f'&begin_day={b_day}&end_day={e_day}&PCount=100')
    req = urllib.request.Request(url, headers={
        'X-Requested-With': 'XMLHttpRequest',
        'Referer': f'{MOPS_BASE}/t51sb10'})
    for attempt in range(3):
        try:
            with opener.open(req, timeout=30) as r:
                return r.read().decode('utf-8', 'ignore')
        except Exception as e:
            if attempt < 2:
                time.sleep(attempt + 1)
            else:
                print(f'  [自結] search failed ({keyword}/{typek}): {e}')
    return None


def _parse_listing(html: str) -> list[dict]:
    """從 t51sb10 結果抽每筆公告 (onclick + h{i}4 主旨)."""
    if not html:
        return []
    items = []
    for m in re.finditer(
        r'seq_no\.value="(\d+)";document\.fm\.spoke_time\.value="(\d+)";'
        r'document\.fm\.spoke_date\.value="(\d+)";document\.fm\.i\.value="(\d+)";'
        r'document\.fm\.co_id\.value="(\d+)";document\.fm\.TYPEK\.value="(\w+)"', html):
        seq_no, spoke_time, spoke_date, idx, co_id, typek = m.groups()
        subj_m = re.search(rf"name='h{idx}4' value='([^']*)'", html)
        subject = (subj_m.group(1) if subj_m else '').replace('\n', ' ').strip()
        items.append({
            'co_id': co_id, 'seq_no': seq_no, 'spoke_date': spoke_date,
            'spoke_time': spoke_time, 'TYPEK': typek, 'subject': subject,
        })
    return items


def _classify(subject: str) -> str:
    if '處置' in subject:
        return '處置股'
    if '注意' in subject:
        return '注意股'
    return '自願自結'


def search_self_reported(date_ad: str, progress: bool = True) -> list[dict]:
    """掃指定日期所有 自結/注意/處置 公告, 去重. 回傳 list[announcement item]."""
    y, m, d = (int(x) for x in date_ad.split('-'))
    roc_year = y - 1911
    opener = _new_session()
    seen = {}
    for keyword in KEYWORDS:
        for kind in MARKET_KINDS:
            html = _search_once(opener, keyword, roc_year, m, d, d, '', kind)
            for it in _parse_listing(html):
                key = (it['co_id'], it['seq_no'], it['spoke_date'], it['spoke_time'])
                if key not in seen:
                    it['source_type'] = _classify(it['subject'])
                    it['opener'] = opener
                    seen[key] = it
            time.sleep(0.3)
    items = list(seen.values())
    if progress:
        by = {}
        for it in items:
            by[it['source_type']] = by.get(it['source_type'], 0) + 1
        print(f'  [自結] {date_ad}: {len(items)} 筆 ({by})')
    return items


# ---------------------------------------------------------------- 明細層
def _fetch_detail_html(item: dict, throttle: float = 0.4) -> Optional[str]:
    opener = item.get('opener') or _new_session()
    sd = item['spoke_date']  # 西元 8 碼 YYYYMMDD
    form = {
        'step': '2', 'firstin': 'true', 'off': '1', 'colorchg': '1',
        'TYPEK': item['TYPEK'], 'co_id': item['co_id'], 'seq_no': item['seq_no'],
        'spoke_date': sd, 'spoke_time': item['spoke_time'],
        'year': str(int(sd[:4]) - 1911), 'month': str(int(sd[4:6])),
    }
    data = urllib.parse.urlencode(form).encode()
    req = urllib.request.Request(f'{MOPS_BASE}/ajax_t05st01', data=data, headers={
        'Content-Type': 'application/x-www-form-urlencoded',
        'X-Requested-With': 'XMLHttpRequest',
        'Referer': f'{MOPS_BASE}/t51sb10'})
    time.sleep(throttle)
    for attempt in range(3):
        try:
            with opener.open(req, timeout=30) as r:
                return r.read().decode('utf-8', 'ignore')
        except Exception:
            if attempt < 2:
                time.sleep(attempt + 1)
    return None


def _strip(html: str) -> str:
    t = re.sub(r'<[^>]+>', ' ', html)
    return re.sub(r'\s+', ' ', t).replace('&nbsp;', ' ').strip()


def _extract_name(html: str) -> Optional[str]:
    m = re.search(r'本資料由\s*[(（][^)）]*[)）]\s*\d+\s+([^\s公]+)', html)
    return m.group(1).strip() if m else None


def _seg_after(text: str, label: str, stops: list[str]) -> Optional[str]:
    """回傳 label 之後到下一個 stop label 之間的字串. 重疊標籤 (互為子字串) 不當停止點."""
    i = text.find(label)
    if i < 0:
        return None
    seg = text[i + len(label):]
    cut = len(seg)
    for sl in stops:
        if sl in label or label in sl:  # 跳過自身與重疊 (合併營業收入淨額 vs 營業收入)
            continue
        j = seg.find(sl)
        if 0 <= j < cut:
            cut = j
    return seg[:cut]


def _data_region(text: str) -> str:
    """裁出「3.財務業務資訊 ~ 4.有無」區段, 避開頁尾「近期營業收入及損益資訊」等雜訊。"""
    start = text.find('財務業務')
    if start < 0:
        start = 0
    end = -1
    for marker in ('4.有', '4、有', '４.有', '有無「'):
        j = text.find(marker, start + 4)
        if j > start:
            end = j if end < 0 else min(end, j)
    return text[start:end] if end > start else text[start:start + 2500]


def parse_template_a(text: str) -> dict:
    """注意/處置股統一表: 抽「最近一月 = 自結數」欄 (toks[0]) + 增減% (toks[1]).
    toks[2] = 最近一季查核數 (= 官方季報, 僅留存比對, 不當主訊號)."""
    region = _data_region(text)
    region_unit = _find_unit(region)  # 有些公司在表頭宣告一次「單位:新台幣佰萬」
    out = {'parse_method': 'regex_A'}
    all_labels = [l for _, ls in A_FIELDS for l in ls]
    for key, labels in A_FIELDS:
        for lab in labels:
            seg = _seg_after(region, lab, all_labels)
            if seg is None:
                continue
            toks = re.findall(NUM_RE, seg)
            if not toks:
                continue
            month_val = _to_float(toks[0])
            month_yoy = _pct(toks[1]) if len(toks) >= 2 else None
            quarter_val = _to_float(toks[2]) if len(toks) >= 3 else None
            if key == 'eps':
                out['eps'] = month_val
                out['eps_yoy'] = month_yoy
                out['eps_quarter'] = quarter_val
            else:
                out[key] = _normalize_amount(month_val, _find_unit(seg) or region_unit)
                out[f'{key}_yoy'] = month_yoy
            break
    return out


def parse_template_b(text: str) -> dict:
    """自願自結: 各家會計科目表, 抽當月 (toks[0]) [+累計 toks[1]]. 單位通常表頭宣告一次."""
    out = {'parse_method': 'regex_B'}
    unit = _find_unit(text)
    terminators = ['備註', '尚未經會計師', '實際數字以', '注意事項']
    all_labels = [l for _, ls in B_FIELDS for l in ls] + terminators
    for key, labels in B_FIELDS:
        for lab in labels:
            seg = _seg_after(text, lab, all_labels)
            if seg is None:
                continue
            toks = re.findall(NUM_RE, seg)
            if not toks:
                continue
            month_val = _to_float(toks[0])
            cumulative = _to_float(toks[1]) if len(toks) >= 2 else None
            if key == 'eps':
                out['eps'] = month_val
                out['eps_cumulative'] = cumulative
            else:
                out[key] = _normalize_amount(month_val, unit)
                out[f'{key}_cumulative'] = _normalize_amount(cumulative, unit)
            break
    return out


def _is_template_a(text: str) -> bool:
    return ('最近一月' in text and '最近一季' in text) or '合併自結數' in text


def _meaningful(parsed: dict) -> bool:
    return any(parsed.get(k) is not None for k in ('eps', 'revenue', 'net_income', 'pretax'))


# ---------------------------------------------------------------- AI fallback
_HAIKU_SYS = """你是台股自結財務資訊抽取器。輸入是一則 MOPS 重大訊息的「說明」全文。

任務: 抽出公司「最近一月 / 當月自結數」那一欄的數字 (不是去年同期、不是查核季、不是累計)。
多數公告會同時列當月、累計、去年同期、最近一季查核數 — 你只取「最近一月(自結)」。
金額單位照原文標示 (仟元 / 百萬 / 元),EPS 單位一律「元」。沒有的欄位填 null。
嚴格依 schema 輸出 JSON,不要 markdown。"""


def extract_with_haiku(client, model: str, subject: str, text: str) -> dict:
    from pydantic import BaseModel, Field

    class SelfReportedExtract(BaseModel):
        unit: Optional[str] = Field(None, description="金額單位: 仟元/百萬/元")
        revenue: Optional[float] = Field(None, description="當月自結營業收入")
        operating_income: Optional[float] = Field(None, description="當月自結營業利益/損益")
        pretax: Optional[float] = Field(None, description="當月自結稅前損益")
        net_income: Optional[float] = Field(None, description="當月自結稅後淨利")
        eps: Optional[float] = Field(None, description="當月自結每股盈餘(元)")

    resp = client.messages.parse(
        model=model, max_tokens=300,
        system=[{'type': 'text', 'text': _HAIKU_SYS, 'cache_control': {'type': 'ephemeral'}}],
        messages=[{'role': 'user', 'content': f'主旨: {subject}\n\n說明:\n{text[:2500]}'}],
        output_format=SelfReportedExtract,
    )
    p = resp.parsed_output
    return {
        'parse_method': 'haiku',
        'revenue': _normalize_amount(p.revenue, p.unit),
        'operating_income': _normalize_amount(p.operating_income, p.unit),
        'pretax': _normalize_amount(p.pretax, p.unit),
        'net_income': _normalize_amount(p.net_income, p.unit),
        'eps': p.eps,
    }


# ---------------------------------------------------------------- 主流程
def fetch_self_reported_for_date(date_ad: str, ai_client=None, model: Optional[str] = None,
                                 derive_eps: bool = True, progress: bool = True) -> list[dict]:
    """抓指定日期所有自結公告並解析. 回傳 list[record]."""
    items = search_self_reported(date_ad, progress=progress)
    records = []
    for i, it in enumerate(items, 1):
        if progress and i % 15 == 0:
            print(f'  [自結] 解析 {i}/{len(items)}...', flush=True)
        html = _fetch_detail_html(it)
        if not html:
            continue
        text = _strip(html)
        if _is_template_a(text):
            # 注意/處置股統一表 — 證交所固定格式, 純 regex 最準
            parsed = parse_template_a(text)
        else:
            # 自願自結 — 各家格式天差地別 (純文字敘述 / 多段表 / 億元為單位),
            # 以 Claude Haiku 抽取為主, regex_B 為無 AI / AI 失敗時的後備.
            parsed = None
            if ai_client is not None:
                try:
                    parsed = extract_with_haiku(ai_client, model or 'claude-haiku-4-5',
                                                it['subject'], text)
                except Exception as e:
                    if progress:
                        print(f'  [自結] {it["co_id"]} Haiku 失敗, 退 regex: {str(e)[:40]}')
            if parsed is None or not _meaningful(parsed):
                parsed = parse_template_b(text)
        if not _meaningful(parsed):
            continue
        # 期間 (從主旨抓「N月」)
        pm = re.search(r'(\d{1,2})\s*月', it['subject'])
        period_month = f'{pm.group(1)}月' if pm else None
        rec = {
            'stock_id': it['co_id'],
            'name': _extract_name(html),
            'announce_date': f'{it["spoke_date"][:4]}-{it["spoke_date"][4:6]}-{it["spoke_date"][6:8]}',
            'announce_time': it['spoke_time'],
            'subject': it['subject'],
            'source_type': it['source_type'],
            'period_month': period_month,
            'eps_source': 'self_reported',
            **parsed,
        }
        # 缺 EPS → 用自結淨利 ÷ 流通股數 自算 (月自結淨利 → 月自結 EPS)
        if derive_eps and rec.get('eps') is None and rec.get('net_income'):
            try:
                from fetch_eps import fetch_shares_outstanding
                shares = fetch_shares_outstanding(it['co_id'])
                if shares and shares > 0:
                    rec['eps'] = round(rec['net_income'] * 1000 / shares, 2)
                    rec['eps_source'] = 'self_reported_computed'
                    rec['eps_shares_used'] = shares
            except Exception:
                pass
        records.append(rec)
    if progress:
        print(f'  [自結] 解析出 {len(records)} 筆有效自結資料')
    return records


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / '.env', override=True)
    import os

    date = sys.argv[1] if len(sys.argv) > 1 else '2026-06-01'
    client = None
    try:
        import anthropic
        if os.environ.get('ANTHROPIC_API_KEY'):
            client = anthropic.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])
    except ImportError:
        pass

    recs = fetch_self_reported_for_date(date, ai_client=client,
                                        model=os.environ.get('SCORE_MODEL', 'claude-haiku-4-5'))
    print(f'\n=== {date} 自結速報 {len(recs)} 筆 ===')
    for r in sorted(recs, key=lambda x: (x['source_type'], x['stock_id'])):
        eps = r.get('eps')
        ni = r.get('net_income')
        rev = r.get('revenue')
        yoy = r.get('eps_yoy')
        yoy_s = f' YoY{yoy*100:+.0f}%' if yoy is not None else ''
        print(f"  {r['stock_id']:>5} {r.get('name') or '':<8} [{r['source_type']}] "
              f"{r.get('period_month') or '?':>4} EPS={eps}{yoy_s} "
              f"淨利={ni} 營收={rev} ({r['parse_method']}/{r['eps_source']})")
