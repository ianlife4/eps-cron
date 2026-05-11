"""
一次性 5/11 MOPS 補抓 — 跑完 bg main.py 後執行.

bg main.py 用 FinMind 抓全市場, 但對「公告當天 MOPS 已 release 但 FinMind 還沒同步」
的個股 (如 8096 擎亞 5/11 18:07 公告) 抓不到. 此腳本:
1. 載入 eps_2026-05-11.json (bg 產出)
2. 找 latest != 2026-03-31 的個股
3. 調 fetch_mops_supplement 補抓
4. 重 analyze + score + build report
5. 推 TG 補抓版

僅本機跑, 不上 GHA cron (生產 cron 已有 MOPS hook).
"""
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent.parent
sys.path.insert(0, str(BASE / 'src'))
sys.stdout.reconfigure(encoding='utf-8')

from dotenv import load_dotenv
load_dotenv(BASE / '.env', override=True)

from compare import analyze_one, freshest_quarter_end, quarter_label, build_first_seen_map
from score import score_rule_based
from report import build_report
from notify import send_message, send_document, send_photo, format_daily_summary
from render_image import render_releases_png
from fetch_mops import fetch_mops_supplement
from fetch_eps import fetch_stock_list

TODAY = '2026-05-11'
SNAPSHOT_DIR = BASE / 'data' / 'snapshots'
REPORTS_DIR = BASE / 'reports'


def get_market_classifier(stock_info: list) -> dict:
    out = {}
    for s in stock_info:
        sid = str(s.get('stock_id', '')).strip()
        if not sid:
            continue
        out[sid] = {'name': s.get('stock_name', ''), 'market': s.get('type', '')}
    return out


def main():
    print(f'=== MOPS 補抓 {TODAY} ===')

    # 1. 載入 bg 跑出的 snapshot
    snap_file = SNAPSHOT_DIR / f'eps_{TODAY}.json'
    if not snap_file.exists():
        print(f'❌ {snap_file} 不存在 — bg main.py 還沒跑完?')
        return 1
    eps_data = json.loads(snap_file.read_text(encoding='utf-8'))
    print(f'載入 snapshot: {snap_file.name} ({len(eps_data)} 檔)')

    freshest_q = freshest_quarter_end(TODAY)
    q_label = quarter_label(freshest_q)
    print(f'當期: {q_label} ({freshest_q})')

    # 2. 找需要補 MOPS 的個股
    needs_mops = set()
    for sid, eps in eps_data.items():
        if not eps:
            needs_mops.add(sid)
            continue
        valid_dates = [d for d, fin in eps.items()
                       if isinstance(fin, dict) and fin.get('eps') is not None]
        if not valid_dates or max(valid_dates) != freshest_q:
            needs_mops.add(sid)
    print(f'\n候選補抓: {len(needs_mops)} 檔 (FinMind 沒收到 {q_label})')

    # 3. 撈 MOPS supplement — 掃今天 + 昨天 (catch 18:00 後 / 21:30 前的公告)
    from datetime import timedelta
    mops_data = {}
    for delta in range(3):
        scan_date = (datetime.strptime(TODAY, '%Y-%m-%d') - timedelta(days=delta)).strftime('%Y-%m-%d')
        day_supplement = fetch_mops_supplement(
            scan_date, needs_mops, quarter=1, throttle_sec=0.4, progress=True
        )
        for sid, d_data in day_supplement.items():
            if sid not in mops_data:
                mops_data[sid] = d_data
    print(f'\n=== MOPS 補回 {len(mops_data)} 檔 ===')
    for sid, d_data in sorted(mops_data.items()):
        for date, fin in d_data.items():
            eps_v = fin.get('eps')
            print(f'  {sid}: {date} EPS={eps_v}')

    if not mops_data:
        print('沒有 MOPS 補抓到的 → 不推播')
        return 0

    # 4. 注入 eps_data (不覆蓋 FinMind 既有)
    for sid, d_data in mops_data.items():
        if sid not in eps_data:
            eps_data[sid] = {}
        for date, fin in d_data.items():
            if date not in eps_data[sid]:
                eps_data[sid][date] = fin
            else:
                if eps_data[sid][date].get('eps') is None and fin.get('eps') is not None:
                    eps_data[sid][date].update(fin)

    # 重存 snapshot
    snap_file.write_text(json.dumps(eps_data, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'  ✓ 重存 snapshot 含 MOPS 補的 {len(mops_data)} 檔')

    # 5. 撈股票對照表
    print('\n撈股票對照表...')
    stock_info = fetch_stock_list('ALL')
    seen = set()
    info_list = []
    for s in stock_info:
        sid = str(s.get('stock_id', '')).strip()
        if not sid or sid in seen:
            continue
        seen.add(sid)
        info_list.append(s)
    classifier = get_market_classifier(info_list)

    # 6. 重新 analyze + score 所有股票 (含 MOPS 補的)
    print('\n重 analyze + score...')
    releases = []
    for sid, eps in eps_data.items():
        if not eps:
            continue
        analysis = analyze_one(sid, eps)
        if not analysis.get('has_data'):
            continue
        info = classifier.get(sid, {})
        analysis['name'] = info.get('name', '')
        analysis['market'] = info.get('market', '')
        # is_new = 今天才出現的 latest_quarter
        analysis['is_new'] = analysis.get('latest_quarter') == q_label
        scoring = score_rule_based(analysis)
        analysis.update(scoring)
        releases.append(analysis)
    print(f'  分析 {len(releases)} 檔')

    # 7. stats
    new_count = sum(1 for r in releases if r.get('is_new'))
    hot_count = sum(1 for r in releases if (r.get('score') or 0) >= 8)
    stats = {
        'total_count': len(releases),
        'new_count': new_count,
        'hot_count': hot_count,
        'watch_count': sum(1 for r in releases if 6 <= (r.get('score') or 0) < 8),
        'warn_count': sum(1 for r in releases if (r.get('score') or 0) <= -4),
    }
    print(f'  statistics: {stats}')

    # 8. first_seen + build report
    first_seen_map = build_first_seen_map(SNAPSHOT_DIR, freshest_q, today_str=TODAY)
    excel_path = REPORTS_DIR / f'eps_daily_{TODAY}.xlsx'
    build_report(TODAY, releases, [], stats, str(excel_path),
                 q_label=q_label, first_seen_map=first_seen_map)
    print(f'  ✓ Excel: {excel_path}')

    # 9. 推 TG: 標題說明這是 MOPS 補抓版
    tg_token = os.environ['TG_BOT_TOKEN']
    chat_id = os.environ['TG_CHAT_ID']

    mops_added = [r for r in releases if r['stock_id'] in mops_data]
    mops_added.sort(key=lambda x: (-(x.get('score') if x.get('score') is not None else -99),
                                    -(x.get('latest_eps') or 0)))

    # intro
    intro_lines = [
        f'<b>🔄 MOPS 補抓 {TODAY}</b>',
        f'<i>FinMind 還沒同步, 直接從 MOPS 抓 {q_label} 公告:</i>',
        '',
    ]
    for r in mops_added[:30]:
        sid = r['stock_id']
        name = r.get('name', '')
        eps_v = r.get('latest_eps')
        yoy_pct = (r.get('yoy') or {}).get('pct')
        yoy_str = f' YoY {yoy_pct*100:+.0f}%' if yoy_pct is not None else ''
        score = r.get('score')
        label = r.get('label', '')
        intro_lines.append(f'  <code>{sid}</code> {name} EPS={eps_v}{yoy_str} {label}')
    if len(mops_added) > 30:
        intro_lines.append(f'  ...+ {len(mops_added) - 30} 檔')
    intro_lines.append('')
    intro_lines.append('<i>表格 PNG + Excel 在下方</i>')

    print('\n推 TG intro...')
    send_message(tg_token, chat_id, '\n'.join(intro_lines))
    time.sleep(1.5)

    # PNG: MOPS 補抓的個股
    if mops_added:
        png_path = REPORTS_DIR / f'mops_supplement_{TODAY}.png'
        render_releases_png(
            mops_added,
            title=f'MOPS 補抓 ({len(mops_added)} 檔, 評分降冪 top 30)',
            out_path=str(png_path),
            date_str=TODAY,
            max_rows=30,
            first_seen_map=first_seen_map,
            subtitle='FinMind 還沒同步, MOPS 直連抓回的個股',
        )
        print(f'推 PNG MOPS 補抓...')
        send_photo(tg_token, chat_id, str(png_path),
                   caption=f'🔄 MOPS 補抓 {len(mops_added)} 檔 (含 FinMind 還沒入庫的)')
        time.sleep(1.5)

    # 完整 Excel (含全市場最新)
    print('推 Excel 全量更新版...')
    send_document(tg_token, chat_id, str(excel_path),
                  caption=f'📊 EPS 日報 {TODAY} 補抓版 (新增 {len(mops_added)} 檔 MOPS)')

    print(f'\n=== 完成 ===')
    print(f'MOPS 補抓 {len(mops_added)} 檔 → TG 推完')
    return 0


if __name__ == '__main__':
    sys.exit(main())
