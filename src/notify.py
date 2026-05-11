"""
Telegram Bot 通知模組
- send_message: 文字訊息 (Markdown)
- send_document: 附 Excel 檔
- send_photo: PNG 圖片
- _post 內建 3 次 retry, 避開 TG 偶發 SSL connection reset (WinError 10054)
"""
import os
import json
import time
import urllib.parse
import urllib.request
import urllib.error
from pathlib import Path

TG_API = 'https://api.telegram.org/bot'
RETRY_COUNT = 3
RETRY_BASE_SLEEP = 3  # seconds; 3, 6, 9


def _build_request(method: str, token: str, data: dict, file_path: str = None,
                   file_field: str = 'document',
                   file_mime: str = 'application/octet-stream'):
    """Construct urllib Request. multipart 當 file_path 不為 None."""
    url = f'{TG_API}{token}/{method}'
    if file_path:
        # multipart/form-data
        boundary = '----eps-cron-boundary-' + os.urandom(8).hex()
        body = b''
        for k, v in data.items():
            body += f'--{boundary}\r\n'.encode()
            body += f'Content-Disposition: form-data; name="{k}"\r\n\r\n'.encode()
            body += f'{v}\r\n'.encode()
        with open(file_path, 'rb') as f:
            file_content = f.read()
        fname = Path(file_path).name
        body += f'--{boundary}\r\n'.encode()
        body += f'Content-Disposition: form-data; name="{file_field}"; filename="{fname}"\r\n'.encode()
        body += f'Content-Type: {file_mime}\r\n\r\n'.encode()
        body += file_content + b'\r\n'
        body += f'--{boundary}--\r\n'.encode()
        return urllib.request.Request(url, data=body, headers={
            'Content-Type': f'multipart/form-data; boundary={boundary}'
        })
    encoded = urllib.parse.urlencode(data).encode()
    return urllib.request.Request(url, data=encoded)


def _post(method: str, token: str, data: dict, file_path: str = None,
          file_field: str = 'document',
          file_mime: str = 'application/octet-stream',
          retries: int = RETRY_COUNT) -> dict:
    """call Telegram API with retry.

    - 處理 TG 偶發 SSL connection reset (WinError 10054 / ConnectionResetError)
    - 處理 4xx/5xx HTTPError (TG 有時回 429 rate limit)
    - file_path 不為 None 時用 multipart upload
    """
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            req = _build_request(method, token, data, file_path, file_field, file_mime)
            with urllib.request.urlopen(req, timeout=60) as r:
                resp = json.loads(r.read())
            # TG 200 OK 但 ok:false (例如檔太大 / chat 不存在) → 不重試
            if not resp.get('ok'):
                print(f'  [TG] {method} returned ok=false: {resp.get("description")}')
            return resp
        except (urllib.error.URLError, ConnectionResetError, TimeoutError) as e:
            last_err = e
            if attempt < retries:
                sleep = RETRY_BASE_SLEEP * attempt
                print(f'  [TG] {method} attempt {attempt}/{retries} failed: '
                      f'{type(e).__name__}: {str(e)[:80]} — retry in {sleep}s')
                time.sleep(sleep)
            else:
                print(f'  [TG] {method} all {retries} attempts failed')
                raise
    raise last_err  # unreachable


def send_message(token: str, chat_id: str, text: str, parse_mode: str = 'HTML') -> dict:
    """送純文字。HTML 比 Markdown 在多 emoji 場景穩定。"""
    return _post('sendMessage', token, {
        'chat_id': chat_id,
        'text': text,
        'parse_mode': parse_mode,
        'disable_web_page_preview': 'true',
    })


def send_document(token: str, chat_id: str, file_path: str, caption: str = '') -> dict:
    """送檔案附件 (Excel)."""
    return _post('sendDocument', token, {
        'chat_id': chat_id,
        'caption': caption,
        'parse_mode': 'HTML',
    }, file_path=file_path, file_field='document',
       file_mime='application/octet-stream')


def send_photo(token: str, chat_id: str, photo_path: str, caption: str = '') -> dict:
    """送 PNG 圖片 (TG 內嵌預覽). caption 上限 1024 chars."""
    if len(caption) > 1024:
        caption = caption[:1020] + '…'
    return _post('sendPhoto', token, {
        'chat_id': chat_id,
        'caption': caption,
        'parse_mode': 'HTML',
    }, file_path=photo_path, file_field='photo', file_mime='image/png')


def _fmt_vs_full_year(r: dict) -> str:
    """組「Q1 EPS vs 去年全年」短文字, 例:
      - 已贏全年:   '25全21.83 (263%)'
      - 轉虧為盈:   '25全-0.5 → 轉盈'
      - 還沒贏全年: '25全0.88 (達 22%)'
      - 缺資料:     ''
    用 latest_quarter 推導去年: e.g. 2026Q1 → 2025 → '25全'
    """
    pf = r.get('prior_year_full')
    ach = r.get('achievement_pct')
    latest = r.get('latest_eps')
    q = r.get('latest_quarter', '')
    # 取去年代號 (2026Q1 → 「25」)
    prior_y_short = ''
    try:
        prior_y_short = str(int(q[:4]) - 1)[-2:]
    except (ValueError, IndexError):
        pass
    prefix = f'{prior_y_short}全' if prior_y_short else '去年全'
    if pf is None:
        return ''
    if ach == 'prior_loss' or (isinstance(pf, (int, float)) and pf <= 0):
        if latest and latest > 0:
            return f'{prefix}{pf}→轉盈'
        return f'{prefix}{pf}'
    if isinstance(ach, (int, float)):
        return f'{prefix}{pf} ({ach*100:.0f}%)'
    return f'{prefix}{pf}'


def _fmt_reasons(reasons, max_items: int = 2) -> str:
    """規則版 reasons 是 list[str], AI 版是 str. 統一輸出 '、' 連接.

    舊版 `'、'.join((reasons or [])[:2])` 對 str 會切前 2 字 → bug.
    對齊 report.py._format_reasons.
    """
    if not reasons:
        return ''
    if isinstance(reasons, str):
        return reasons
    if isinstance(reasons, list):
        return '、'.join(str(x) for x in reasons[:max_items])
    return str(reasons)


def format_daily_summary(date_str: str, new_releases: list, top_winners: list,
                         monthly_rev_highlights: list = None) -> str:
    """組裝詳細版摘要訊息 (HTML)。"""
    lines = []
    lines.append(f'<b>📊 EPS 日報 {date_str}</b>')
    lines.append('━' * 12)

    if new_releases:
        lines.append(f'<b>🆕 今日新公告：{len(new_releases)} 檔</b>')
        for r in new_releases[:10]:
            sid = r['stock_id']
            name = r.get('name', '')
            q = r.get('latest_quarter', '')
            eps = r.get('latest_eps')
            score = r.get('score')
            label = r.get('label', '')
            yoy_pct = (r.get('yoy') or {}).get('pct')
            yoy_str = f' YoY {yoy_pct*100:+.0f}%' if yoy_pct is not None else ''
            vs_full = _fmt_vs_full_year(r)
            vs_str = f' | {vs_full}' if vs_full else ''
            lines.append(f'  <code>{sid}</code> {name} {q} EPS={eps}{yoy_str}{vs_str} {label}')
        if len(new_releases) > 10:
            lines.append(f'  ...+ {len(new_releases) - 10} 檔（看 Excel）')
        lines.append('')

    if top_winners:
        lines.append(f'<b>🏆 評分 ≥ 8 高度超預期：{len(top_winners)} 檔</b>')
        for r in top_winners[:5]:
            sid = r['stock_id']
            name = r.get('name', '')
            label = r.get('label', '')
            reasons = _fmt_reasons(r.get('reasons'), max_items=2)
            vs_full = _fmt_vs_full_year(r)
            vs_str = f'｜{vs_full}' if vs_full else ''
            lines.append(f'  <code>{sid}</code> {name} {label}{vs_str} — {reasons}')
        lines.append('')

    # 新增區塊: Q1 一季賺贏全年的明星標的 (按倍數降冪 top 10)
    won_full_year = [
        r for r in new_releases
        if isinstance(r.get('achievement_pct'), (int, float)) and r['achievement_pct'] >= 1.0
    ]
    won_full_year.sort(key=lambda x: -(x.get('achievement_pct') or 0))
    if won_full_year:
        # 推導去年年度: 2026Q1 → 2025
        q0 = (new_releases[0].get('latest_quarter') or '')
        prior_year = ''
        try:
            prior_year = str(int(q0[:4]) - 1)
        except (ValueError, IndexError):
            prior_year = '去年'
        lines.append(f'<b>📈 一季賺贏 {prior_year} 全年：'
                     f'{len(won_full_year)} 檔 (按倍數)</b>')
        for r in won_full_year[:10]:
            sid = r['stock_id']
            name = r.get('name', '')
            eps = r.get('latest_eps')
            pf = r.get('prior_year_full')
            ach = r.get('achievement_pct')
            label = r.get('label', '')
            if isinstance(ach, (int, float)) and pf:
                lines.append(f'  <code>{sid}</code> {name} {eps} / {prior_year}全年{pf} = <b>{ach:.2f}x</b> {label}')
        if len(won_full_year) > 10:
            lines.append(f'  ...+ {len(won_full_year) - 10} 檔')
        lines.append('')

    if monthly_rev_highlights:
        lines.append(f'<b>💰 月營收 YoY ≥ +50%：{len(monthly_rev_highlights)} 檔</b>')
        for r in monthly_rev_highlights[:5]:
            sid = r['stock_id']
            name = r.get('name', '')
            yoy = r.get('rev_yoy', 0)
            lines.append(f'  <code>{sid}</code> {name} {yoy:+.0f}%')
        lines.append('')

    lines.append('📸 <i>表格 PNG 在下方圖片，📥 完整 Excel 在最下方</i>')
    lines.append('<i>(同日公告同色帶, Excel 各分頁含「公告日期」欄)</i>')
    return '\n'.join(lines)


if __name__ == '__main__':
    # 自測：發測試訊息
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / '.env')
    token = os.environ['TG_BOT_TOKEN']
    chat_id = os.environ['TG_CHAT_ID']

    # 模擬資料
    new = [
        {'stock_id': '5386', 'name': '青雲', 'latest_quarter': '2026Q1', 'latest_eps': 43.05,
         'yoy': {'pct': 32.9}, 'label': '🔥 +9'},
        {'stock_id': '2408', 'name': '南亞科', 'latest_quarter': '2026Q1', 'latest_eps': 8.41,
         'yoy': {'pct': 13.34}, 'label': '🔥 +9'},
    ]
    winners = new
    msg = format_daily_summary('2026-05-06 (週三)', new, winners)
    print('Sending test message...')
    r = send_message(token, chat_id, msg)
    print(json.dumps(r, ensure_ascii=False, indent=2))
