"""
Telegram Bot 通知模組
- send_message: 文字訊息 (Markdown)
- send_document: 附 Excel 檔
"""
import os
import json
import urllib.parse
import urllib.request
from pathlib import Path

TG_API = 'https://api.telegram.org/bot'


def _post(method: str, token: str, data: dict, file_path: str = None,
          file_field: str = 'document', file_mime: str = 'application/octet-stream') -> dict:
    """call Telegram API. file_path 不為 None 時用 multipart upload.

    file_field: 'document' (sendDocument) 或 'photo' (sendPhoto)
    """
    url = f'{TG_API}{token}/{method}'
    if file_path:
        # multipart/form-data
        boundary = '----eps-cron-boundary-' + os.urandom(8).hex()
        body = b''
        for k, v in data.items():
            body += f'--{boundary}\r\n'.encode()
            body += f'Content-Disposition: form-data; name="{k}"\r\n\r\n'.encode()
            body += f'{v}\r\n'.encode()
        # file
        with open(file_path, 'rb') as f:
            file_content = f.read()
        fname = Path(file_path).name
        body += f'--{boundary}\r\n'.encode()
        body += f'Content-Disposition: form-data; name="{file_field}"; filename="{fname}"\r\n'.encode()
        body += f'Content-Type: {file_mime}\r\n\r\n'.encode()
        body += file_content + b'\r\n'
        body += f'--{boundary}--\r\n'.encode()
        req = urllib.request.Request(url, data=body, headers={
            'Content-Type': f'multipart/form-data; boundary={boundary}'
        })
    else:
        encoded = urllib.parse.urlencode(data).encode()
        req = urllib.request.Request(url, data=encoded)
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read())


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
            lines.append(f'  <code>{sid}</code> {name} {q} EPS={eps}{yoy_str} {label}')
        if len(new_releases) > 10:
            lines.append(f'  ...+ {len(new_releases) - 10} 檔（看 Excel）')
        lines.append('')

    if top_winners:
        lines.append(f'<b>🏆 評分 ≥ 8 高度超預期：{len(top_winners)} 檔</b>')
        for r in top_winners[:5]:
            sid = r['stock_id']
            name = r.get('name', '')
            label = r.get('label', '')
            reasons = '、'.join((r.get('reasons') or [])[:2])
            lines.append(f'  <code>{sid}</code> {name} {label} — {reasons}')
        lines.append('')

    if monthly_rev_highlights:
        lines.append(f'<b>💰 月營收 YoY ≥ +50%：{len(monthly_rev_highlights)} 檔</b>')
        for r in monthly_rev_highlights[:5]:
            sid = r['stock_id']
            name = r.get('name', '')
            yoy = r.get('rev_yoy', 0)
            lines.append(f'  <code>{sid}</code> {name} {yoy:+.0f}%')
        lines.append('')

    lines.append('📥 <i>完整 Excel 在下面 ↓</i>')
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
