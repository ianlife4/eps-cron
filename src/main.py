"""
主控：每日 EPS 監控流程
1. 撈全市場股票清單
2. 抓 EPS / 月營收
3. 比對昨日 snapshot 找新公告
4. 計算 YoY / 累計達成率 / 評分
5. 產出 Excel
6. 推 Telegram
"""
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# Set up paths
BASE = Path(__file__).parent.parent
sys.path.insert(0, str(BASE / 'src'))

from dotenv import load_dotenv
load_dotenv(BASE / '.env', override=True)

from fetch_eps import (fetch_stock_list, fetch_quarterly_eps, fetch_batch,
                       save_snapshot, load_snapshot)
from fetch_revenue import fetch_monthly_revenue, fetch_batch_monthly, analyze_revenue
from compare import analyze_one
from score import score_rule_based
from report import build_report
from notify import send_message, send_document, format_daily_summary

# AI 評分 (可選, 沒 API key 自動 fallback 到規則版)
try:
    from score_ai import score_one as ai_score_one
    import anthropic
    AI_AVAILABLE = True
except ImportError:
    AI_AVAILABLE = False


SNAPSHOT_DIR = BASE / 'data' / 'snapshots'
HISTORICAL_DIR = BASE / 'data' / 'historical'
REPORTS_DIR = BASE / 'reports'


def get_market_classifier(stock_info: list) -> dict:
    """{stock_id: {name, market}} 對照表."""
    out = {}
    for s in stock_info:
        sid = str(s.get('stock_id', '')).strip()
        if not sid: continue
        out[sid] = {
            'name': s.get('stock_name', ''),
            'market': s.get('type', ''),
        }
    return out


def is_target_market(market: str) -> bool:
    """三層分流: 上市/上櫃/創新板每天追"""
    return market in ('twse', 'tpex', 'TWSE', 'TPEx', 'tib', 'TIB')


def detect_new_releases(today_eps: dict, yesterday_eps: dict) -> set:
    """偵測哪些股票今天有新季度公告 (相對昨日 snapshot)。
    回傳 set of stock_id。
    """
    new = set()
    for sid, today in today_eps.items():
        if not today: continue
        yesterday = yesterday_eps.get(sid, {})
        # 找今天有但昨天沒有的季度（且 EPS 不為 None）
        for date, fin in today.items():
            if fin.get('eps') is None: continue
            if date not in yesterday or yesterday[date].get('eps') is None:
                new.add(sid)
                break
    return new


def run_daily(force_all: bool = False, scope: str = 'twse_tpex',
              start_date: str = '2024-01-01', no_ai: bool = False,
              max_ai: int = 100, no_tg: bool = False, no_revenue: bool = False):
    """主流程
    scope: 'twse_tpex' (預設) / 'all' / 'test' (5 檔測試)
    force_all: True 表示全部當作新公告 (首次跑 / backfill)
    no_ai: 不用 AI 評分 (純規則)
    max_ai: AI 評分最大筆數 (依規則版預分數排序取 top N)
    no_tg: 不發 TG (本機測試用)
    no_revenue: 不抓月營收 (省時間)
    """
    today_str = datetime.now().strftime('%Y-%m-%d')
    end_date = datetime.now().strftime('%Y-%m-%d')
    print(f'=== EPS Daily Run {today_str} ===')

    # 1. 撈股票清單
    print('[1] 撈股票清單...')
    if scope == 'test':
        stock_info = [
            {'stock_id': '5386', 'stock_name': '青雲', 'type': 'tpex'},
            {'stock_id': '2408', 'stock_name': '南亞科', 'type': 'twse'},
            {'stock_id': '2330', 'stock_name': '台積電', 'type': 'twse'},
            {'stock_id': '6640', 'stock_name': '均華', 'type': 'tpex'},
            {'stock_id': '3324', 'stock_name': '雙鴻', 'type': 'twse'},
        ]
    else:
        all_stocks = fetch_stock_list('ALL')
        # 去重 (FinMind 同代號可能多筆)
        seen = set()
        stock_info = []
        for s in all_stocks:
            sid = str(s.get('stock_id', '')).strip()
            if not sid or sid in seen: continue
            seen.add(sid)
            stock_info.append(s)

    classifier = get_market_classifier(stock_info)

    # 篩選目標市場
    if scope == 'twse_tpex':
        target_ids = [s['stock_id'] for s in stock_info if is_target_market(s.get('type', ''))]
    else:
        target_ids = [s['stock_id'] for s in stock_info]
    # 過濾 ETF (代號 0050~0099 之類，及 00xxxx)
    target_ids = [sid for sid in target_ids if not sid.startswith('00') and len(sid) == 4 and sid.isdigit()]

    print(f'  目標股票數: {len(target_ids)}')

    # 2. 載入昨日 snapshot
    yesterday_snap_files = sorted(SNAPSHOT_DIR.glob('eps_*.json'), reverse=True)
    yesterday_eps = {}
    if yesterday_snap_files and not force_all:
        latest = yesterday_snap_files[0]
        if latest.stem != f'eps_{today_str}':
            yesterday_eps = json.loads(latest.read_text(encoding='utf-8'))
            print(f'  載入 snapshot: {latest.name} ({len(yesterday_eps)} 檔)')

    # 3. 抓今日 EPS (有 FinMind token 加速)
    has_token = bool(os.environ.get('FINMIND_TOKEN'))
    throttle = 0.6 if not has_token else 0.7  # token: 6000/hr ≈ 1.67/sec, 留餘量 0.7s/req
    print(f'[2] 抓 EPS 資料 (throttle={throttle}s, token={"✓" if has_token else "✗"})...')
    eps_data = fetch_batch(target_ids, start_date, end_date, throttle_sec=throttle, progress=True)

    # 存 snapshot
    snap_file = save_snapshot(eps_data, f'eps_{today_str}')
    print(f'  存 snapshot: {snap_file.name}')

    # 4. 偵測新公告
    if force_all:
        new_set = set(target_ids)
        print(f'[3] FORCE_ALL: 全部 {len(new_set)} 檔當新公告處理')
    else:
        new_set = detect_new_releases(eps_data, yesterday_eps)
        print(f'[3] 偵測到 {len(new_set)} 檔有新公告')

    # 5. 分析 + 規則版評分 (全部)
    print('[4] 分析 + 規則版評分...')
    releases = []
    for sid in target_ids:
        if sid not in eps_data or not eps_data[sid]:
            continue
        analysis = analyze_one(sid, eps_data[sid])
        if not analysis.get('has_data'):
            continue
        info = classifier.get(sid, {})
        analysis['name'] = info.get('name', '')
        analysis['market'] = info.get('market', '')
        analysis['is_new'] = sid in new_set
        scoring = score_rule_based(analysis)
        analysis.update(scoring)
        releases.append(analysis)
    print(f'  分析 {len(releases)} 檔')

    # 5b. AI 評分 — 只跑「新公告 + 規則版預分數高/低」，避免 backfill 燒錢
    use_ai = (not no_ai) and AI_AVAILABLE and os.environ.get('ANTHROPIC_API_KEY')
    if use_ai:
        ai_client = anthropic.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])
        ai_model = os.environ.get('SCORE_MODEL', 'claude-haiku-4-5')
        # 篩選 AI 評分對象: 新公告中, 依規則版分數絕對值排序取 top N
        candidates = [r for r in releases if r.get('is_new')]
        candidates.sort(key=lambda x: -abs(x.get('score') or 0))
        targets = candidates[:max_ai]
        print(f'[4b] AI 評分 ({ai_model}, 最多 {max_ai} 檔)... 共 {len(targets)} 檔候選')
        ai_calls = ai_fail = 0
        for r in targets:
            try:
                ai_scoring = ai_score_one(ai_client, r, model=ai_model)
                # AI 評分蓋過規則版
                r.update(ai_scoring)
                ai_calls += 1
            except Exception as e:
                print(f'  ⚠️ {r["stock_id"]} AI fallback: {str(e)[:60]}', flush=True)
                ai_fail += 1
            if ai_calls % 20 == 0 and ai_calls > 0:
                print(f'  [AI {ai_calls}/{len(targets)}]...', flush=True)
        print(f'  AI 評分完成: {ai_calls} 筆 (失敗 fallback {ai_fail})')
    else:
        print('[4b] AI 評分跳過 (--no-ai 或無 API key)')

    # 統計
    stats = {
        'total_count': len(releases),
        'new_count': sum(1 for r in releases if r.get('is_new')),
        'hot_count': sum(1 for r in releases if (r.get('score') or 0) >= 8),
        'watch_count': sum(1 for r in releases if 6 <= (r.get('score') or 0) < 8),
        'warn_count': sum(1 for r in releases if (r.get('score') or 0) <= -4),
    }
    print(f'  統計: {stats}')

    # 5c. 月營收 (依 user 需求 Q9=A 合併進同一份報告)
    monthly_data = []
    if not no_revenue:
        print('[4c] 抓月營收 + 算 YoY...')
        rev_start = (datetime.now().replace(month=1, day=1).year - 1)
        rev_start_date = f'{rev_start}-01-01'
        rev_raw = fetch_batch_monthly(target_ids, rev_start_date, end_date,
                                      throttle_sec=throttle, progress=True)
        # 偵測「月營收新公告」: 比對昨日月份是否變化 (簡化: 只要有最新月就算)
        for sid in target_ids:
            if sid not in rev_raw or not rev_raw[sid]:
                continue
            ra = analyze_revenue(rev_raw[sid])
            if not ra.get('has_data'): continue
            info = classifier.get(sid, {})
            ra['stock_id'] = sid
            ra['name'] = info.get('name', '')
            ra['market'] = info.get('market', '')
            # 改名以對齊 report.write_revenue 的 schema
            ra['ym'] = ra.pop('latest_ym')
            ra['yoy'] = (ra.pop('mom_yoy_pct') or 0) * 100 if ra.get('mom_yoy_pct') is not None else None
            ra['accumulated'] = ra.pop('accum_rev')
            ra['accum_yoy'] = (ra.pop('accum_yoy_pct') or 0) * 100 if ra.get('accum_yoy_pct') is not None else None
            monthly_data.append(ra)
        print(f'  月營收 {len(monthly_data)} 檔')
    else:
        print('[4c] --no-revenue, 跳過月營收')

    # 6. 產 Excel
    print('[5] 產出 Excel...')
    excel_path = REPORTS_DIR / f'eps_daily_{today_str}.xlsx'
    build_report(today_str, releases, monthly_data, stats, str(excel_path))
    print(f'  產出: {excel_path}')

    # 7. 發 Telegram (僅當有新公告)
    if no_tg:
        print('[6] --no-tg, 跳過 Telegram')
    elif stats['new_count'] > 0 or force_all:
        print('[6] 推 Telegram...')
        token = os.environ['TG_BOT_TOKEN']
        chat_id = os.environ['TG_CHAT_ID']

        new_only = sorted([r for r in releases if r.get('is_new')],
                          key=lambda x: -(x.get('score') or -99))
        winners = [r for r in new_only if (r.get('score') or 0) >= 8]
        msg = format_daily_summary(today_str, new_only[:30], winners)
        send_message(token, chat_id, msg)
        time.sleep(1)
        send_document(token, chat_id, str(excel_path),
                      caption=f'📊 EPS 日報 {today_str} ({stats["new_count"]} 檔新公告)')
        print('  ✓ TG 推播完成')
    else:
        print('[6] 無新公告，跳過 TG 推播')

    print(f'\n=== 完成 ===')
    return {'stats': stats, 'excel': str(excel_path), 'releases_count': len(releases)}


if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('--force-all', action='store_true', help='全部當新公告 (首次/backfill)')
    p.add_argument('--scope', default='twse_tpex', choices=['twse_tpex', 'all', 'test'])
    p.add_argument('--start', default='2024-01-01', help='抓資料起始日')
    p.add_argument('--no-ai', action='store_true', help='不用 AI 評分')
    p.add_argument('--max-ai', type=int, default=100, help='AI 評分最大筆數 (預設 100)')
    p.add_argument('--no-tg', action='store_true', help='不發 Telegram (本機測試)')
    p.add_argument('--no-revenue', action='store_true', help='不抓月營收')
    args = p.parse_args()
    run_daily(force_all=args.force_all, scope=args.scope, start_date=args.start,
              no_ai=args.no_ai, max_ai=args.max_ai, no_tg=args.no_tg,
              no_revenue=args.no_revenue)
