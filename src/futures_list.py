"""TAIFEX 個股期貨標的清單 — 用來在自結速報標記「有股期」的個股。

有股期 = 可做空/加槓桿, 較好操作, 值得在報表標出來。
來源: https://www.taifex.com.tw/cht/2/stockLists (同「開低走高掃描」scan.py)。
抓不到就回空 set (不標記, 不影響主流程)。
"""
import re
import urllib.request
from typing import Optional

_URL = 'https://www.taifex.com.tw/cht/2/stockLists'
_CACHE: Optional[set] = None


def fetch_stock_futures_set(force: bool = False) -> set:
    """回傳「有個股期貨」的股票/ETF 代號 set。per-run cache; 失敗回空 set。"""
    global _CACHE
    if _CACHE is not None and not force:
        return _CACHE
    codes = set()
    try:
        req = urllib.request.Request(_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=30) as r:
            html = r.read().decode('utf-8', 'replace')
        for tr in re.findall(r'<tr[\s>].*?</tr>', html, re.S):
            cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', tr, re.S)
            cells = [re.sub(r'<[^>]+>', '', c).replace('\r', '').replace('\t', '')
                     .replace('\n', '').strip() for c in cells]
            if len(cells) < 11:
                continue
            fut_code, code, is_fut = cells[0], cells[2], cells[4]
            if (re.fullmatch(r'[0-9A-Z]{2,3}', fut_code)
                    and re.fullmatch(r'\d{4,6}', code) and is_fut):
                codes.add(code)
    except Exception as e:
        print(f'  [股期清單] 抓取失敗 (不標記): {e}')
    _CACHE = codes
    return codes


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    s = fetch_stock_futures_set()
    print(f'TAIFEX 有股期的標的: {len(s)} 檔')
    for c in ['3042', '4904', '2330', '8222', '1718', '2434', '9999']:
        print(f'  {c}: {"有股期 ✔" if c in s else "無"}')
