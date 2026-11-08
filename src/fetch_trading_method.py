"""
變更交易方法(全額交割) / 恢復交易方法 追蹤 — 財報惡化的領先訊號 (2026-11-08 上線)

兩條來源, 互為備援:
 A. MOPS 重大訊息 (ajax_t51sb10 主旨全文檢索, 上市 KIND=L / 上櫃 KIND=O)
    公司依規定公告「列為 / 併案 / 繼續 變更交易方法」「加採 / 取消 分盤」「恢復普通交易方法」。
    主旨常被換行切斷 (「變更交易 方法」「變更 交易方法」) → 多組短關鍵字取聯集。
    內文是證交所/櫃買統一模板:
      1.公告處置之日期 2.引用之條款及發生緣由 3.處理結果 4.股票開始(併案)變更交易方法…之日期
      恢復版: 1.公告股票回復原狀日期 2.…發生緣由 3.公告內容 4.股票回復原狀之日期
    → 抽 生效日 / 原因標籤 / 處理結果。公告日通常比生效日早 2~4 個交易日 (領先官方名單)。
 B. 官方「目前」名單 (OpenAPI, 免 token):
      上市 TWSE  /v1/exchangeReport/TWT85U   (PeriodicCallAuctionTrading == '**' → 分盤)
      上櫃 TPEx  /openapi/v1/tpex_cmode      (變更交易/分盤/管理股票/停止交易 四旗標, 全形「Ｙ」)
    每天存 data/trading_method/official_list.json, 跟上一份比對 → 名單新增/移出/分盤變動。
    有對應 MOPS 公告的 = 生效 (不重推); 沒有的 = MOPS 漏抓, 以「名單變動」補推。
    某市場當天抓失敗 → 該市場不比對 (避免假的「移出」)。

上市/上櫃規則不同 (2026-08-17 查證, 2911 統一實業 / 4530 內文實證):
  上市 營業細則§49: 淨值 < 股本 1/2 → 變更交易方法; 恢復需「最近二期」淨值均 ≥ 股本 1/2。
                    淨值 < 股本 3/10 等 → 併採分盤集合競價 (§49-2)。
  上櫃 業務規則§12: 淨值 < 股本 1/2 → 變更 (I 1); < 3/10 → 加採分盤 (V 2);
                    恢復只要「最近一期」≥ 股本 1/2 且較前期增加 (較易脫身)。
  共同觸發: 逾期未申報財報 / 會計師保留、無法表示、否定意見 / 繼續經營疑慮 / 退票、債務違約 / 重整 …

狀態檔 (GHA 每天 commit):
  data/trading_method/events.json        — 所有看過的 MOPS 事件 (key 去重, 判斷「今天新的」)
  data/trading_method/official_list.json — 上一份官方名單
"""
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent))
from fetch_mops import _new_session, PAT_OVERRUN  # noqa: E402
import fetch_self_reported as _fsr  # noqa: E402  (共用 t51sb10 搜尋/明細/去 tag 工具)

STATE_DIR = Path(os.environ.get('TM_STATE_DIR') or
                 Path(__file__).parent.parent / 'data' / 'trading_method')
EVENTS_FILE = 'events.json'
LIST_FILE = 'official_list.json'

# 主旨被換行切斷 → 短關鍵字取聯集 (2026 實測: 2314「變更交易 方法」只有「分盤」「變更交易」抓得到)
KEYWORDS = ['交易方法', '變更交易', '恢復交易', '分盤', '全額交割']
MARKET_KINDS = {'L': '上市', 'O': '上櫃'}   # 興櫃無全額交割制度
SCAN_DAYS = 6          # 每天回掃 6 天 (週末/連假/前一晚 19:30 後公告都補得到), 已看過的靠 events.json 去重
HISTORY_DAYS = 180     # Excel 歷史分頁顯示範圍
RECENT_MATCH_DAYS = 20  # 名單變動 ↔ MOPS 公告 對得上的天數

TWSE_URL = 'https://openapi.twse.com.tw/v1/exchangeReport/TWT85U'
TPEX_URL = 'https://www.tpex.org.tw/openapi/v1/tpex_cmode'

# 事件 → (嚴重度排序, 色系)。色系給 PNG/Excel/TG 共用
EVENT_STYLE = {
    '新列全額+分盤': (0, 'deep'),
    '併案變更+分盤': (0, 'deep'),
    '續列全額+分盤': (1, 'deep'),
    '停止買賣':      (1, 'deep'),
    '新列全額':      (2, 'red'),
    '併案變更':      (2, 'red'),
    '名單新增':      (2, 'red'),
    '分盤啟用':      (2, 'red'),
    '續列全額':      (3, 'orange'),
    '恢復買賣(仍全額)': (3, 'orange'),
    '取消分盤':      (4, 'yellow'),
    '分盤取消':      (4, 'yellow'),
    '恢復普通交易':  (5, 'green'),
    '名單移出':      (6, 'gray'),   # 沒對應恢復公告 → 可能是終止上市櫃, 不給綠色
    '其他':          (6, 'gray'),
}
TG_ICON = {'deep': '⛔', 'red': '🔴', 'orange': '🟠', 'yellow': '🟡', 'green': '🟢', 'gray': '⚪'}
RECOVER_RULE = {
    '上市': '恢復:最近二期淨值均≥股本1/2',
    '上櫃': '恢復:最近一期淨值≥股本1/2且較前期增加',
}


# ---------------------------------------------------------------- 小工具
def _roc_to_iso(s: str) -> Optional[str]:
    """'115/04/07' / '115年04月07日' → '2026-04-07'."""
    m = re.search(r'(\d{2,3})\s*[/年.-]\s*(\d{1,2})\s*[/月.-]\s*(\d{1,2})', s or '')
    if not m:
        return None
    y, mo, d = int(m.group(1)) + 1911, int(m.group(2)), int(m.group(3))
    try:
        return datetime(y, mo, d).strftime('%Y-%m-%d')
    except ValueError:
        return None


def _iso8(d8: str) -> str:
    """'20260401' → '2026-04-01'."""
    return f'{d8[:4]}-{d8[4:6]}-{d8[6:8]}' if len(d8) == 8 else d8


def _days_between(a: str, b: str) -> int:
    return abs((datetime.strptime(a, '%Y-%m-%d') - datetime.strptime(b, '%Y-%m-%d')).days)


def event_key(e: dict) -> str:
    return f"{e['stock_id']}_{e['spoke_date']}_{e['spoke_time']}_{e['seq_no']}"


def event_color(event: str) -> str:
    return EVENT_STYLE.get(event, (6, 'gray'))[1]


def event_rank(event: str) -> int:
    return EVENT_STYLE.get(event, (6, 'gray'))[0]


# ---------------------------------------------------------------- MOPS 搜尋 / 明細
def _search(opener, kw: str, roc_year: int, month: int, b: int, e: int, kind: str) -> Optional[str]:
    for attempt in range(3):
        html = _fsr._search_once(opener, kw, roc_year, month, b, e, '', kind)
        if html and PAT_OVERRUN.search(html):
            wait = 10 * (attempt + 1)
            print(f'  [交易方法] MOPS 限流 Overrun → 等 {wait}s')
            time.sleep(wait)
            continue
        return html
    return None


def _month_spans(start: datetime, end: datetime):
    """[start, end] 切成 (年, 月, 起日, 迄日) — t51sb10 一次只能查同一個月。"""
    cur = start
    while cur <= end:
        nxt = (cur.replace(day=28) + timedelta(days=4)).replace(day=1)
        last = min(end, nxt - timedelta(days=1))
        yield cur.year, cur.month, cur.day, last.day
        cur = nxt


def search_announcements(start: str, end: str, throttle: float = 0.9) -> list[dict]:
    """掃 [start, end] 的交易方法相關公告 (上市+上櫃, 多關鍵字聯集去重)。"""
    s = datetime.strptime(start, '%Y-%m-%d')
    e = datetime.strptime(end, '%Y-%m-%d')
    opener = _new_session()
    seen = {}
    for y, m, b, d in _month_spans(s, e):
        for kind in MARKET_KINDS:
            for kw in KEYWORDS:
                html = _search(opener, kw, y - 1911, m, b, d, kind)
                for it in _fsr._parse_listing(html or ''):
                    k = (it['co_id'], it['seq_no'], it['spoke_date'], it['spoke_time'])
                    if k not in seen:
                        it['kind'] = kind
                        it['opener'] = opener
                        seen[k] = it
                time.sleep(throttle)
    return list(seen.values())


def _fetch_detail_text(item: dict, throttle: float = 0.9) -> tuple[str, Optional[str]]:
    for attempt in range(3):
        html = _fsr._fetch_detail_html(item, throttle=throttle)
        if html and PAT_OVERRUN.search(html):
            time.sleep(10 * (attempt + 1))
            continue
        if not html:
            return '', None
        return _fsr._strip(html), _fsr._extract_name(html)
    return '', None


# ---------------------------------------------------------------- 解析
_NEG = ('已無', '已不低於', '已大於', '已達', '已符合', '不再', '原因業已消滅', '已消滅', '已不')


def _reason_tags(reason: str) -> list[str]:
    """發生緣由 → 短標籤。先丟掉「已無…/已不低於…」這類改善子句, 避免把舊原因當新原因。"""
    clauses = [c for c in re.split(r'[，,。；;]', reason or '') if not any(n in c for n in _NEG)]
    t = ''.join(clauses)
    tags = []
    # 有些公告只寫條號 (2380 虹光「同細則第49條第1項第1款…」):
    #   上市 營業細則§49 I 1 / 上櫃 業務規則§12 I 1 = 淨值低於股本 1/2; 上櫃 §12 V 2 = 低於 3/10
    half = re.search(r'第\s*49\s*條第\s*1\s*項第\s*1\s*款|第\s*12\s*條第\s*1\s*項第\s*1\s*款', t)
    tenth = re.search(r'第\s*12\s*條第\s*1\s*項第\s*1\s*款及第\s*5\s*項第\s*2\s*款|第\s*5\s*項第\s*2\s*款', t)
    if '二分之一' in t or '1/2' in t or half:
        tags.append('淨值低於股本1/2')
    if '十分之三' in t or '3/10' in t or tenth:
        tags.append('淨值低於股本3/10')
    if '無法表示意見' in t or '否定意見' in t:
        tags.append('會計師無法表示/否定意見')
    elif '保留意見' in t:
        tags.append('會計師保留意見')
    if '繼續經營' in t and '不確定' in t:
        tags.append('繼續經營疑慮')
    if ('期限' in t or '逾期' in t) and ('財務報告' in t or '財報' in t) and ('未' in t):
        tags.append('逾期未申報財報')
    if '獨立董事' in t:
        tags.append('獨董不足')
    if '退票' in t or '無法如期償還' in t or '債務' in t or '違約' in t:
        tags.append('退票/債務違約')
    if '重整' in t:
        tags.append('聲請重整')
    if '經營權' in t and '營業範圍' in t:
        tags.append('經營權異動+營業範圍變更')
    if not tags:
        # 認不得的條款 (如 4744 皇將 櫃買§12 I 14) → 至少標條號, 原文在 Excel「發生緣由」
        m = re.search(r'第\s*(\d+)\s*條(?:之\s*(\d+))?\s*第\s*(\d+)\s*項第\s*(\d+)\s*款', t)
        if m:
            art = m.group(1) + (f'-{m.group(2)}' if m.group(2) else '')
            tags.append(f'§{art}第{m.group(3)}項第{m.group(4)}款(見原文)')
    return tags


def classify_event(subject: str, result: str, text: str) -> str:
    """主旨 + 處理結果 + 內文 → 事件類型 (EVENT_STYLE 的 key)."""
    s = re.sub(r'\s+', '', f'{subject} {result}')
    still_full = '繼續列為變更交易' in s or '仍繼續列為變更' in s or '繼續採預收款券' in s
    if ('恢復' in s or '回復原狀' in text) and still_full:
        # 5481 (2025-09): 停止買賣原因消滅 → 恢復買賣, 但「繼續列為變更交易方法」— 不是脫離全額!
        return '恢復買賣(仍全額)'
    if '恢復' in s or '回復原狀' in text:
        return '恢復普通交易'
    if '取消分盤' in s:
        return '取消分盤'
    periodic = '分盤' in s or '每30分鐘撮合' in text.replace(' ', '')
    if '停止買賣' in result and '變更交易' not in result:
        return '停止買賣'
    if '併案' in s:
        base = '併案變更'
    elif '繼續' in s or '維持' in s:
        base = '續列全額'
    elif '變更交易' in s or '全額交割' in s:
        base = '新列全額'
    else:
        return '其他'
    return f'{base}+分盤' if periodic else base


def _seg(text: str, pat: str) -> str:
    m = re.search(pat, text)
    return re.sub(r'\s+', '', m.group(1)) if m else ''


def _is_bond_only(subject: str) -> bool:
    """公司債(含可轉債)被列變更交易方法 — 不是股票本身。
    1589 永冠 2026-02-23「…第四次無擔保轉換公司債(永冠四KY 15894)列為變更交易方法有價證券…」
    (主旨後段也有「有價證券」, 所以要看有沒有「上市/上櫃有價證券」「股票」「普通股」)。"""
    return '公司債' in subject and not re.search(r'(上市|上櫃)有價證券|股票|普通股', subject)


def parse_detail(item: dict, text: str, name: Optional[str]) -> Optional[dict]:
    """MOPS 明細 → 事件 dict。澄清/說明類 (沒有統一模板) 回 None。"""
    subject = re.sub(r'\s+', '', item.get('subject', ''))
    if _is_bond_only(subject):
        return None
    is_template = ('處置之日期' in text) or ('回復原狀' in text)
    if not is_template:
        # 「經證交所同意自X日起恢復交易」= 重大訊息暫停交易後恢復 (1218 泰山/1909 榮成), 跟全額交割無關;
        # 「…澄清/說明」是公司回應新聞。只留主旨明講交易方法/全額交割/分盤的。
        if '澄清' in subject or '說明' in subject:
            return None
        # 「變更交易方式」(5481/2067) 也算; 但「變更交易對象」(1904 出售房地) 不算
        if not any(k in subject for k in ('交易方法', '交易方式', '全額交割', '分盤')):
            return None
    reason = _seg(text, r'2\.[^:：]{0,40}發生緣由\s*[:：]\s*(.+?)\s*3\.(?:處理結果|公告內容)')
    result = (_seg(text, r'3\.處理結果[^:：]*[:：]\s*(.+?)\s*4\.')
              or _seg(text, r'3\.公告內容\s*[:：]\s*(.+?)\s*4\.'))
    eff = _roc_to_iso(_seg(text, r'4\.(?:股票開始|股票回復原狀)[^:：]*[:：]\s*(\d{2,3}/\d{1,2}/\d{1,2})'))
    if not eff:   # 主旨「自民國115年04月07日起」
        m = re.search(r'自民國(\d{2,3}年\d{1,2}月\d{1,2}日)', subject)
        eff = _roc_to_iso(m.group(1)) if m else None
    event = classify_event(subject, result, text)
    # 市場別: 明細頁「本資料由 (上市公司) 1435 中福 公司提供」最準; 退而求其次用搜尋時的 KIND
    m = re.search(r'本資料由\s*[(（](上市|上櫃)公司', text)
    market = m.group(1) if m else MARKET_KINDS.get(item.get('kind'), '')
    tags = _reason_tags(reason) if event not in ('恢復普通交易', '取消分盤', '恢復買賣(仍全額)') else []
    if event == '恢復買賣(仍全額)':
        tags = ['停止買賣原因消滅,仍列全額']
    elif event == '恢復普通交易':
        tags = ['淨值回升≥股本1/2'] if ('二分之一' in reason or '1/2' in reason) else ['原因消滅']
    elif event == '取消分盤':
        tags = ['淨值回升≥股本3/10'] if '十分之三' in reason else ['分盤原因消滅']
    return {
        'stock_id': item['co_id'],
        'name': name or '',
        'market': market,
        'event': event,
        'is_correction': '更正' in subject,
        'announce_date': _iso8(item['spoke_date']),
        'announce_time': item['spoke_time'],
        'effective_date': eff or '',
        'reason_tags': tags,
        'reason': reason[:300],
        'result': result[:120],
        'subject': subject[:120],
        'spoke_date': item['spoke_date'], 'spoke_time': item['spoke_time'], 'seq_no': item['seq_no'],
        'source': 'MOPS',
    }


def fetch_events(start: str, end: str, known_keys: set = frozenset(),
                 throttle: float = 0.9, progress: bool = True) -> list[dict]:
    """抓 [start, end] 的交易方法事件; known_keys 內的不再抓明細 (省 MOPS 請求)。"""
    items = search_announcements(start, end, throttle=throttle)
    out = []
    skipped = 0
    for it in items:
        k = f"{it['co_id']}_{it['spoke_date']}_{it['spoke_time']}_{it['seq_no']}"
        if k in known_keys:
            continue
        text, name = _fetch_detail_text(it, throttle=throttle)
        ev = parse_detail(it, text, name) if text else None
        if ev is None:
            skipped += 1
            continue
        out.append(ev)
    if progress:
        print(f'  [交易方法] {start}~{end}: 公告 {len(items)} 則, 新解析 {len(out)} 則'
              f'{f", 略過澄清/無模板 {skipped}" if skipped else ""}')
    return out


# ---------------------------------------------------------------- 官方名單
_Y = ('Y', 'Ｙ', 'y')


def _get_json(url: str, timeout: int = 30):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 eps-cron',
                                               'Accept': 'application/json'})
    last = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.loads(r.read().decode('utf-8-sig'))
        except Exception as e:
            last = e
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f'{url}: {last}')


def fetch_official_lists() -> dict:
    """{'twse_ok', 'tpex_ok', 'stocks': {code: {name, market, altered, periodic, managed, suspended, freq}}}"""
    out = {'twse_ok': False, 'tpex_ok': False, 'stocks': {}, 'fetched_at': datetime.now().isoformat(timespec='seconds')}
    try:
        for r in _get_json(TWSE_URL):
            code = str(r.get('Code', '')).strip()
            if not code:
                continue
            out['stocks'][code] = {
                'name': str(r.get('Name', '')).strip(), 'market': '上市',
                'altered': True, 'periodic': '*' in str(r.get('PeriodicCallAuctionTrading', '')),
                'managed': False, 'suspended': False, 'freq': '',
            }
        out['twse_ok'] = True
    except Exception as e:
        print(f'  [交易方法] TWSE 官方名單失敗 (上市不比對): {e}')
    try:
        for r in _get_json(TPEX_URL):
            r = {k.strip(): (v.strip() if isinstance(v, str) else v) for k, v in r.items()}
            code = str(r.get('SecuritiesCompanyCode', '')).strip()
            if not code:
                continue
            out['stocks'][code] = {
                'name': r.get('CompanyName', ''), 'market': '上櫃',
                'altered': r.get('AlteredTrading') in _Y,
                'periodic': r.get('PeriodicTrading') in _Y,
                'managed': r.get('ManagedStock') in _Y,
                'suspended': r.get('SuspensionOfTrading') in _Y,
                'freq': r.get('MatchingFrequency', ''),
            }
        out['tpex_ok'] = True
    except Exception as e:
        print(f'  [交易方法] TPEx 官方名單失敗 (上櫃不比對): {e}')
    return out


def diff_official(prev: Optional[dict], cur: dict) -> list[dict]:
    """名單比對 → 名單新增 / 名單移出 / 分盤啟用 / 分盤取消 (只比兩邊都抓成功的市場)。"""
    if not prev:
        return []
    ok = {'上市': prev.get('twse_ok') and cur.get('twse_ok'),
          '上櫃': prev.get('tpex_ok') and cur.get('tpex_ok')}
    p, c = prev.get('stocks', {}), cur.get('stocks', {})
    changes = []
    for code in sorted(set(p) | set(c)):
        a, b = p.get(code), c.get(code)
        market = (b or a)['market']
        if not ok.get(market):
            continue
        was = bool(a and a.get('altered'))
        now = bool(b and b.get('altered'))
        name = (b or a).get('name', '')
        if now and not was:
            changes.append({'stock_id': code, 'name': name, 'market': market, 'event': '名單新增'})
        elif was and not now:
            changes.append({'stock_id': code, 'name': name, 'market': market, 'event': '名單移出'})
        if a and b and a.get('periodic') != b.get('periodic'):
            changes.append({'stock_id': code, 'name': name, 'market': market,
                            'event': '分盤啟用' if b.get('periodic') else '分盤取消'})
    return changes


# ---------------------------------------------------------------- 狀態檔
def _load(name: str, default):
    f = STATE_DIR / name
    if f.exists():
        try:
            return json.loads(f.read_text(encoding='utf-8'))
        except Exception as e:
            print(f'  [交易方法] 讀 {f.name} 失敗 ({e}), 當作空的')
    return default


def save_state(tm: dict):
    """push 完再存 (萬一推播前就掛, 下一輪還會當新事件重推, 不會漏)。"""
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    events = sorted(tm['_all_events'].values(), key=lambda e: (e['announce_date'], e['announce_time']))
    (STATE_DIR / EVENTS_FILE).write_text(json.dumps(events, ensure_ascii=False, indent=1), encoding='utf-8')
    off = tm.get('official') or {}
    prev = tm.get('_prev_official') or {}
    # 某市場今天抓失敗 → 保留上一份該市場資料, 明天仍能比對
    keep = dict(off)
    if not off.get('twse_ok') or not off.get('tpex_ok'):
        stocks = dict(off.get('stocks', {}))
        for code, s in prev.get('stocks', {}).items():
            if (s['market'] == '上市' and not off.get('twse_ok')) or (s['market'] == '上櫃' and not off.get('tpex_ok')):
                stocks[code] = s
        keep['stocks'] = stocks
        keep['twse_ok'] = off.get('twse_ok') or prev.get('twse_ok', False)
        keep['tpex_ok'] = off.get('tpex_ok') or prev.get('tpex_ok', False)
    keep['date'] = tm['date']
    (STATE_DIR / LIST_FILE).write_text(json.dumps(keep, ensure_ascii=False, indent=1), encoding='utf-8')


# ---------------------------------------------------------------- 主入口
def scan_trading_method(today: str, name_map: Optional[dict] = None,
                        futures_set: Optional[set] = None, progress: bool = True) -> dict:
    """每日掃描。回傳:
      alerts     — 今天要推的 (新 MOPS 事件 + 沒對應公告的名單變動), 嚴重度排序
      list_changes — 官方名單變動 (含已有公告對應的 = 生效)
      history    — 近 HISTORY_DAYS 天全部 MOPS 事件 (Excel)
      official   — 今天官方名單 (Excel), 每檔附最近一則 MOPS 事件
    狀態檔要在推播後呼叫 save_state(tm) 才寫。
    """
    name_map = name_map or {}
    futures_set = futures_set or set()
    old_events = _load(EVENTS_FILE, [])
    all_events = {event_key(e): e for e in old_events}
    start = (datetime.strptime(today, '%Y-%m-%d') - timedelta(days=SCAN_DAYS)).strftime('%Y-%m-%d')
    new = fetch_events(start, today, known_keys=set(all_events), progress=progress)

    def _family(ev):  # 更正前後類型可能不同 (2314 台揚: 併案變更+分盤 → 更正為 續列全額+分盤)
        return 'recover' if ev in ('恢復普通交易', '取消分盤') else 'worse'

    alerts = []
    # 同一輪同股同方向多則 (原公告+更正) → 只推最晚那則
    latest_new = {}
    for e in new:
        k = (e['stock_id'], _family(e['event']))
        if k not in latest_new or (e['announce_date'], e['announce_time']) > \
                (latest_new[k]['announce_date'], latest_new[k]['announce_time']):
            latest_new[k] = e
    old_snapshot = list(all_events.values())
    for e in new:
        e['first_seen'] = today
        if not e['name']:
            e['name'] = name_map.get(e['stock_id'], '')
        all_events[event_key(e)] = e
        if latest_new[(e['stock_id'], _family(e['event']))] is not e:
            continue
        # 前幾輪已推過同股同方向事件, 10 天內的「更正」→ 只記錄不重推
        if e['is_correction'] and any(
                o['stock_id'] == e['stock_id'] and _family(o['event']) == _family(e['event'])
                and _days_between(o['announce_date'], e['announce_date']) <= 10
                for o in old_snapshot):
            continue
        alerts.append(dict(e))

    prev_official = _load(LIST_FILE, None)
    official = fetch_official_lists()
    changes = diff_official(prev_official, official)
    # 名單變動有對應的近期 MOPS 事件 = 生效, 不重推; 沒有 = MOPS 漏抓 → 補推
    for ch in changes:
        want = ('恢復普通交易',) if ch['event'] in ('名單移出',) else \
               ('取消分盤',) if ch['event'] == '分盤取消' else None
        related = [o for o in all_events.values() if o['stock_id'] == ch['stock_id']
                   and _days_between(o['announce_date'], today) <= RECENT_MATCH_DAYS
                   and (want is None and o['event'] not in ('恢復普通交易',)
                        or want is not None and o['event'] in want)]
        ch['matched'] = bool(related)
        ch['announce_date'] = ''
        ch['effective_date'] = today
        if related:
            ch['reason_tags'] = ['官方名單已生效']
        elif ch['event'] == '名單移出':
            ch['reason_tags'] = ['MOPS未見恢復公告,可能終止上市櫃']
        else:
            ch['reason_tags'] = ['官方名單變動,MOPS未見公告']
        ch['source'] = '官方名單'
        if not related:
            alerts.append(dict(ch))

    # 官方名單每檔附最近一則 MOPS 事件
    latest = {}
    for e in all_events.values():
        if e['stock_id'] not in latest or (e['announce_date'], e['announce_time']) > \
                (latest[e['stock_id']]['announce_date'], latest[e['stock_id']]['announce_time']):
            latest[e['stock_id']] = e
    official_rows = []
    for code, s in official.get('stocks', {}).items():
        le = latest.get(code)
        official_rows.append({
            'stock_id': code, 'name': s['name'] or name_map.get(code, ''), 'market': s['market'],
            'altered': s['altered'], 'periodic': s['periodic'], 'managed': s['managed'],
            'suspended': s['suspended'], 'freq': s.get('freq', ''),
            'last_event': le['event'] if le else '', 'last_date': le['announce_date'] if le else '',
            'last_reason': '、'.join(le['reason_tags']) if le else '',
            'recover_rule': RECOVER_RULE.get(s['market'], '') if s['altered'] else '',
        })

    cutoff = (datetime.strptime(today, '%Y-%m-%d') - timedelta(days=HISTORY_DAYS)).strftime('%Y-%m-%d')
    history = [e for e in all_events.values() if e['announce_date'] >= cutoff]
    for rows in (alerts, history, official_rows, changes):
        for r in rows:
            r['has_futures'] = r['stock_id'] in futures_set
    alerts.sort(key=lambda r: (event_rank(r['event']), r['stock_id']))
    history.sort(key=lambda r: (r['announce_date'], r['announce_time']), reverse=True)
    official_rows.sort(key=lambda r: (r['market'], not r['altered'], not r['periodic'], r['stock_id']))
    if progress:
        n_alt = sum(1 for r in official_rows if r['altered'])
        print(f'  [交易方法] 新事件 {len(new)} / 名單變動 {len(changes)} / 要推 {len(alerts)}; '
              f'官方名單 全額 {n_alt} 檔 (上市 {"✓" if official["twse_ok"] else "✗"} '
              f'上櫃 {"✓" if official["tpex_ok"] else "✗"})')
    return {
        'date': today, 'alerts': alerts, 'list_changes': changes, 'history': history,
        'official': official, 'official_rows': official_rows,
        '_all_events': all_events, '_prev_official': prev_official,
    }


def backfill(since: str, until: str):
    """一次性回補 events.json (歷史事件標 first_seen=公告日, 不會被當成新事件推)。"""
    all_events = {event_key(e): e for e in _load(EVENTS_FILE, [])}
    evs = fetch_events(since, until, known_keys=set(all_events))
    for e in evs:
        e['first_seen'] = e['announce_date']
        all_events[event_key(e)] = e
    official = fetch_official_lists()
    save_state({'date': until, '_all_events': all_events, 'official': official, '_prev_official': None})
    print(f'回補完成: 事件 {len(all_events)} 則 → {STATE_DIR / EVENTS_FILE}')


if __name__ == '__main__':
    import argparse
    sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser()
    ap.add_argument('--backfill', help='回補起日 YYYY-MM-DD (到今天)')
    ap.add_argument('--scan', help='模擬某天掃描 (不寫狀態檔)')
    a = ap.parse_args()
    if a.backfill:
        backfill(a.backfill, datetime.now().strftime('%Y-%m-%d'))
    else:
        tm = scan_trading_method(a.scan or datetime.now().strftime('%Y-%m-%d'))
        for r in tm['alerts']:
            print(r['stock_id'], r['name'], r['market'], r['event'], r.get('effective_date'),
                  '、'.join(r.get('reason_tags') or []))
