"""
PIL-based PNG 表格渲染 — 用於 Telegram 內嵌預覽
- 不依賴 Excel/LibreOffice (cron 環境用 GHA Ubuntu 也能跑)
- 跨平台字型 fallback: Windows JhengHei → Linux Noto CJK
"""
import os
import sys
from pathlib import Path
from typing import Optional

from PIL import Image, ImageDraw, ImageFont


# === 字型解析 (Windows + Linux + macOS fallback) ===
_FONT_CANDIDATES = [
    # Windows
    'C:/Windows/Fonts/msjh.ttc',         # Microsoft JhengHei
    'C:/Windows/Fonts/msjhbd.ttc',       # JhengHei Bold (only used for bold lookup)
    'C:/Windows/Fonts/mingliu.ttc',      # 細明體
    'C:/Windows/Fonts/simsun.ttc',       # 新細明體 fallback
    # Linux (GHA Ubuntu 用 fonts-noto-cjk)
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc',
    '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
    # macOS
    '/System/Library/Fonts/PingFang.ttc',
    '/Library/Fonts/Microsoft/Microsoft JhengHei.ttf',
]


def _find_font(bold: bool = False) -> Optional[str]:
    """挑第一個存在的中文字型 (有 bold 變體優先給 bold 請求)。"""
    bold_keywords = ('Bold', 'bold', 'msjhbd', 'NotoSansCJK-Bold')
    if bold:
        for p in _FONT_CANDIDATES:
            if any(k in p for k in bold_keywords) and Path(p).exists():
                return p
    for p in _FONT_CANDIDATES:
        if Path(p).exists():
            return p
    return None


def _load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    fpath = _find_font(bold=bold) or _find_font(bold=False)
    if fpath:
        try:
            return ImageFont.truetype(fpath, size=size)
        except Exception:
            pass
    return ImageFont.load_default()


# === 配色 (對齊 Excel v3 風格) ===
COL_TITLE_BG = (192, 0, 0)         # 紅 (主標題)
COL_HEADER_BG = (47, 84, 150)      # 深藍 (header 列)
COL_WHITE = (255, 255, 255)
COL_BORDER = (191, 191, 191)
COL_BG_HOT = (255, 217, 102)       # 鮮黃 — score≥8
COL_BG_WARM = (248, 203, 173)      # 暖橘 — score≥4
COL_BG_COLD = (180, 199, 231)      # 冷藍 — score≤-4
COL_BG_TODAY = (152, 240, 152)     # 鮮綠 — 今日剛公告 (蓋過 score 色, 更醒目)
COL_BG_ROW_A = (255, 255, 255)     # 白底
COL_BG_ROW_B = (242, 242, 242)     # 淺灰
COL_TEXT = (0, 0, 0)
COL_TEXT_HIGHLIGHT = (156, 0, 6)   # 深紅 (鮮黃底字)
COL_TEXT_TODAY = (0, 102, 0)       # 深綠 (鮮綠底字, 對比清晰)
COL_FOOTER = (128, 128, 128)


def _row_bg(score) -> tuple:
    if isinstance(score, (int, float)):
        if score >= 8:
            return COL_BG_HOT
        if score >= 4:
            return COL_BG_WARM
        if score <= -4:
            return COL_BG_COLD
    return COL_BG_ROW_A


def _row_text_color(score) -> tuple:
    if isinstance(score, (int, float)) and score >= 8:
        return COL_TEXT_HIGHLIGHT
    return COL_TEXT


def _strip_unrenderable(s: str) -> str:
    """移除 CJK 字型不支援的 emoji / 特殊符號 (避免變方塊)."""
    if not s:
        return s
    out = []
    for ch in s:
        cp = ord(ch)
        # emoji broad ranges (cover 🆕 🏆 🔥 📊 📑 🎯 ⚠️ 等)
        if 0x1F100 <= cp <= 0x1FAFF: continue   # supplement + symbols + emoticons
        if 0x2600 <= cp <= 0x27BF:  continue   # dingbats / misc symbols
        if cp in (0x2728, 0xFE0F, 0x200D):     continue  # variation selectors
        # 比較符號 ≥ ≤ 換成 ASCII
        if cp == 0x2265: out.append('>='); continue
        if cp == 0x2264: out.append('<='); continue
        out.append(ch)
    # 清掉因為移除 emoji 留下的多餘空白
    return ' '.join(''.join(out).split()).strip()


def render_releases_png(releases: list, title: str, out_path: str,
                        date_str: str = '', max_rows: int = 30,
                        first_seen_map: dict = None,
                        subtitle: str = '',
                        today_str: str = '',
                        sort_by_date: bool = False) -> str:
    """渲染 releases 列表為 PNG, 傳回檔案路徑。

    欄位順序: 代號 / 名稱 / 季 / EPS / 去年同季 / YoY% / 去年全年 / 達成率 / 倍數 / QoQ% / 評分 / 級別 / 評分理由 / 公告日期

    sort_by_date:
      - True (用於「今日新公告」): 日期 desc → score desc → EPS desc
        最新公告在最上面, 今日的所有 row 自然集中於頂端
      - False (用於「已贏全年」等): score desc → EPS desc (本來邏輯)

    今日剛公告 (first_seen == today_str) 的列用鮮綠突顯, 蓋過 score 色.
    today_str 未傳時自動推導 (date_str 或 first_seen_map 中最大日期).
    """
    has_date = bool(first_seen_map)

    if sort_by_date and has_date:
        # 三層: 日期 desc → score desc → EPS desc
        rows = sorted(
            releases,
            key=lambda x: (
                first_seen_map.get(x.get('stock_id'), '0000-00-00'),
                x.get('score') if x.get('score') is not None else -99,
                x.get('latest_eps') or 0,
            ),
            reverse=True,
        )[:max_rows]
    else:
        # 兩層: score desc → EPS desc
        rows = sorted(
            releases,
            key=lambda x: (-(x.get('score') if x.get('score') is not None else -99),
                           -(x.get('latest_eps') or 0)),
        )[:max_rows]
    # 推導「今日」: 優先參數, 退到 date_str, 再退到 first_seen_map 最大值
    if not today_str:
        today_str = date_str
    if not today_str and first_seen_map:
        valid = [d for d in first_seen_map.values() if d and d != '—']
        if valid:
            today_str = max(valid)

    cols = [
        ('代號', 70, 'center'),
        ('名稱', 105, 'center'),
        ('季', 75, 'center'),
        ('EPS', 70, 'center'),
        ('去年同季', 80, 'center'),
        ('YoY%', 85, 'center'),
        ('去年全年', 85, 'center'),  # NEW: Q1 vs 去年全年 EPS
        ('達成率', 78, 'center'),
        ('倍數', 65, 'center'),      # NEW: Q1 / 去年全年, 直觀
        ('QoQ%', 80, 'center'),
        ('評分', 55, 'center'),
        ('級別', 95, 'center'),
        ('評分理由', 320, 'left'),
    ]
    if has_date:
        cols.append(('公告日期', 100, 'center'))

    # === Layout 量測 ===
    pad_x = 16
    title_h = 56
    sub_h = 28 if (date_str or subtitle) else 0
    header_h = 36
    row_h = 30
    foot_h = 28

    table_w = sum(c[1] for c in cols)
    canvas_w = table_w + pad_x * 2
    canvas_h = title_h + sub_h + header_h + row_h * len(rows) + foot_h + pad_x

    img = Image.new('RGB', (canvas_w, canvas_h), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    # 字型
    f_title = _load_font(20, bold=True)
    f_sub = _load_font(13)
    f_header = _load_font(14, bold=True)
    f_body = _load_font(13)
    f_body_bold = _load_font(13, bold=True)
    f_foot = _load_font(11)

    # === 標題列 ===
    draw.rectangle((0, 0, canvas_w, title_h), fill=COL_TITLE_BG)
    title_clean = _strip_unrenderable(title)
    title_w = draw.textlength(title_clean, font=f_title)
    draw.text(((canvas_w - title_w) / 2, (title_h - 22) / 2),
              title_clean, fill=COL_WHITE, font=f_title)

    # 副標題
    y = title_h
    if sub_h:
        sub_text = subtitle or f'資料日期 {date_str} | FinMind + Claude Haiku 4.5 評分'
        sub_text = _strip_unrenderable(sub_text)
        draw.rectangle((0, y, canvas_w, y + sub_h), fill=(222, 235, 247))
        sub_w = draw.textlength(sub_text, font=f_sub)
        draw.text(((canvas_w - sub_w) / 2, y + 6), sub_text,
                  fill=(31, 56, 100), font=f_sub)
        y += sub_h

    # === Header 列 ===
    x = pad_x
    draw.rectangle((pad_x, y, pad_x + table_w, y + header_h), fill=COL_HEADER_BG)
    for label, w, _align in cols:
        tw = draw.textlength(label, font=f_header)
        draw.text((x + (w - tw) / 2, y + (header_h - 16) / 2),
                  label, fill=COL_WHITE, font=f_header)
        x += w
    # header 邊框
    draw.line([(pad_x, y), (pad_x + table_w, y)], fill=COL_BORDER, width=1)
    draw.line([(pad_x, y + header_h), (pad_x + table_w, y + header_h)],
              fill=COL_BORDER, width=1)
    y += header_h

    # === 資料列 ===
    for r in rows:
        sc = r.get('score')
        # 今日剛公告 → 鮮綠蓋過 score 色 (最醒目)
        sid_first_seen = first_seen_map.get(r.get('stock_id')) if has_date else None
        is_today = bool(today_str and sid_first_seen == today_str)
        if is_today:
            bg = COL_BG_TODAY
            fg = COL_TEXT_TODAY
            font_used = f_body_bold
        else:
            bg = _row_bg(sc)
            fg = _row_text_color(sc)
            font_used = f_body_bold if bg == COL_BG_HOT else f_body

        draw.rectangle((pad_x, y, pad_x + table_w, y + row_h), fill=bg)

        yoy_pct = (r.get('yoy') or {}).get('pct')
        yoy_str = f'{yoy_pct*100:+.0f}%' if yoy_pct is not None else '—'
        qoq_pct = (r.get('qoq') or {}).get('pct')
        qoq_str = f'{qoq_pct*100:+.0f}%' if qoq_pct is not None else '—'
        ach = r.get('achievement_pct')
        pf = r.get('prior_year_full')
        if isinstance(ach, (int, float)):
            ach_str = f'{ach*100:.0f}%'
            multi_str = f'{ach:.2f}x'
        elif ach == 'prior_loss':
            ach_str = '虧轉盈'
            multi_str = '虧轉盈'
        else:
            ach_str = '—'
            multi_str = '—'
        pf_str = f'{pf}' if pf is not None else '—'
        reasons = r.get('reasons')
        if isinstance(reasons, list):
            reasons_str = '、'.join(str(x) for x in reasons[:2])
        elif isinstance(reasons, str):
            reasons_str = reasons
        else:
            reasons_str = ''
        date_str_val = ''
        if has_date:
            date_str_val = first_seen_map.get(r.get('stock_id'), '—')

        values = [
            str(r.get('stock_id', '')),
            (r.get('name') or '')[:6],
            r.get('latest_quarter', ''),
            f'{r.get("latest_eps")}' if r.get('latest_eps') is not None else '—',
            f'{r.get("yoy_eps")}' if r.get('yoy_eps') is not None else '—',
            yoy_str,
            pf_str,        # 去年全年 EPS (新)
            ach_str,
            multi_str,     # 倍數 (新)
            qoq_str,
            f'{sc}' if sc is not None else '—',
            r.get('level', ''),
            reasons_str,
        ]
        if has_date:
            values.append(date_str_val)

        x = pad_x
        for (col_label, w, align), val in zip(cols, values):
            txt = _strip_unrenderable(str(val))
            # 裁切過長文字
            max_chars = max(1, int(w / 14))
            if len(txt) > max_chars and col_label == '評分理由':
                txt = txt[:max_chars] + '…'
            tw = draw.textlength(txt, font=font_used)
            if align == 'center':
                tx = x + (w - tw) / 2
            elif align == 'right':
                tx = x + w - tw - 6
            else:
                tx = x + 8
            ty = y + (row_h - 16) / 2
            draw.text((tx, ty), txt, fill=fg, font=font_used)
            # 直線分隔
            draw.line([(x, y), (x, y + row_h)], fill=COL_BORDER, width=1)
            x += w
        # 收尾右邊線
        draw.line([(pad_x + table_w, y), (pad_x + table_w, y + row_h)],
                  fill=COL_BORDER, width=1)
        # 底線
        draw.line([(pad_x, y + row_h), (pad_x + table_w, y + row_h)],
                  fill=COL_BORDER, width=1)
        y += row_h

    # === Footer ===
    y += 6
    today_hint = f' / 今日 {today_str} 公告 (鮮綠)' if today_str else ''
    foot_text = (f'共 {len(releases)} 檔 (顯示 top {len(rows)}).  '
                 f'高度超預期 >=8 (鮮黃) / 有亮點 >=4 (暖橘) / 衰退 <=-4 (冷藍){today_hint}')
    foot_text = _strip_unrenderable(foot_text)
    draw.text((pad_x, y), foot_text, fill=COL_FOOTER, font=f_foot)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, format='PNG', optimize=True)
    return out_path


def render_self_reported_png(records: list, title: str, out_path: str,
                             date_str: str = '', max_rows: int = 30,
                             subtitle: str = '') -> str:
    """渲染自結速報為 PNG (TG 內嵌預覽)。

    欄位: 代號 / 名稱 / 類別 / 期間 / 自結EPS / YoY% / 評分 / 級別 / 評分理由
    依評分上色 (同 EPS 表): >=8 鮮黃 / >=4 暖橘 / <=-4 冷藍。依評分降冪排序。
    自結EPS 帶 * = 自結淨利÷股數自算; 無自結EPS (純營收) 不評分、殿後。
    """
    rows = sorted(records,
                  key=lambda r: (r.get('score') if r.get('score') is not None else -99,
                                 r.get('eps') if r.get('eps') is not None else -999),
                  reverse=True)[:max_rows]

    cols = [
        ('代號', 62, 'center'), ('名稱', 92, 'center'), ('類別', 72, 'center'),
        ('期間', 50, 'center'), ('自結EPS', 80, 'center'), ('YoY%', 76, 'center'),
        ('評分', 56, 'center'), ('級別', 96, 'center'), ('評分理由', 300, 'left'),
    ]
    pad_x, title_h, sub_h, header_h, row_h, foot_h = 16, 56, 28, 36, 30, 28
    table_w = sum(c[1] for c in cols)
    canvas_w = table_w + pad_x * 2
    canvas_h = title_h + sub_h + header_h + row_h * max(1, len(rows)) + foot_h + pad_x

    img = Image.new('RGB', (canvas_w, canvas_h), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    f_title = _load_font(20, bold=True)
    f_sub = _load_font(13)
    f_header = _load_font(14, bold=True)
    f_body = _load_font(13)
    f_body_bold = _load_font(13, bold=True)
    f_foot = _load_font(11)

    draw.rectangle((0, 0, canvas_w, title_h), fill=COL_TITLE_BG)
    tc = _strip_unrenderable(title)
    tw = draw.textlength(tc, font=f_title)
    draw.text(((canvas_w - tw) / 2, (title_h - 22) / 2), tc, fill=COL_WHITE, font=f_title)

    y = title_h
    sub_text = _strip_unrenderable(subtitle or f'資料日期 {date_str} | MOPS 自結數 + Claude 評分，領先官方財報 (月/累計，非季報)')
    draw.rectangle((0, y, canvas_w, y + sub_h), fill=(222, 235, 247))
    sw = draw.textlength(sub_text, font=f_sub)
    draw.text(((canvas_w - sw) / 2, y + 6), sub_text, fill=(31, 56, 100), font=f_sub)
    y += sub_h

    x = pad_x
    draw.rectangle((pad_x, y, pad_x + table_w, y + header_h), fill=COL_HEADER_BG)
    for label, w, _a in cols:
        lw = draw.textlength(label, font=f_header)
        draw.text((x + (w - lw) / 2, y + (header_h - 16) / 2), label, fill=COL_WHITE, font=f_header)
        x += w
    draw.line([(pad_x, y), (pad_x + table_w, y)], fill=COL_BORDER, width=1)
    draw.line([(pad_x, y + header_h), (pad_x + table_w, y + header_h)], fill=COL_BORDER, width=1)
    y += header_h

    for r in rows:
        sc = r.get('score')
        bg = _row_bg(sc)
        fg = _row_text_color(sc)
        font_used = f_body_bold if bg == COL_BG_HOT else f_body
        draw.rectangle((pad_x, y, pad_x + table_w, y + row_h), fill=bg)
        eps = r.get('eps')
        if eps is None:
            eps_s = '—'
        else:
            eps_s = f'{eps}*' if str(r.get('eps_source', '')).endswith('computed') else f'{eps}'
        yoy = r.get('eps_yoy')
        yoy_s = f'{yoy * 100:+.0f}%' if yoy is not None else '—'
        reasons = r.get('reasons')
        reasons_s = reasons if isinstance(reasons, str) else (
            '、'.join(str(x) for x in reasons[:2]) if isinstance(reasons, list) else '')
        values = [
            str(r.get('stock_id', '')), (r.get('name') or '')[:6],
            r.get('source_type', ''), r.get('period_month') or '—',
            eps_s, yoy_s,
            f'{sc}' if sc is not None else '—',
            r.get('level') or '—',
            reasons_s,
        ]
        x = pad_x
        for (col_label, w, align), val in zip(cols, values):
            txt = _strip_unrenderable(str(val))
            max_chars = max(1, int(w / 14))
            if len(txt) > max_chars and col_label == '評分理由':
                txt = txt[:max_chars] + '…'
            vw = draw.textlength(txt, font=font_used)
            tx = x + (w - vw) / 2 if align == 'center' else (x + w - vw - 6 if align == 'right' else x + 8)
            draw.text((tx, y + (row_h - 16) / 2), txt, fill=fg, font=font_used)
            draw.line([(x, y), (x, y + row_h)], fill=COL_BORDER, width=1)
            x += w
        draw.line([(pad_x + table_w, y), (pad_x + table_w, y + row_h)], fill=COL_BORDER, width=1)
        draw.line([(pad_x, y + row_h), (pad_x + table_w, y + row_h)], fill=COL_BORDER, width=1)
        y += row_h

    y += 6
    n_att = sum(1 for r in records if r.get('source_type') == '注意股')
    n_dis = sum(1 for r in records if r.get('source_type') == '處置股')
    n_vol = sum(1 for r in records if r.get('source_type') == '自願自結')
    foot = (f'共 {len(records)} 檔 (注意股 {n_att} / 處置股 {n_dis} / 自願自結 {n_vol}, 顯示 top {len(rows)}).  '
            f'評分: >=8 鮮黃 / >=4 暖橘 / <=-4 冷藍.  * = 自結淨利/股數自算')
    draw.text((pad_x, y), _strip_unrenderable(foot), fill=COL_FOOTER, font=f_foot)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, format='PNG', optimize=True)
    return out_path


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    sample = [
        {'stock_id': '5289', 'name': '宜鼎', 'latest_quarter': '2026Q1', 'latest_eps': 57.49,
         'yoy_eps': 3.68, 'yoy': {'pct': 14.62}, 'qoq': {'pct': 5.21},
         'achievement_pct': 2.634, 'score': 9, 'level': '高度超預期',
         'reasons': 'YoY爆增1462%、Q1一季賺贏全年2.6倍、QoQ+521%'},
        {'stock_id': '4973', 'name': '廣穎', 'latest_quarter': '2026Q1', 'latest_eps': 7.46,
         'yoy_eps': -0.03, 'yoy': {'pct': 249.66}, 'qoq': {'pct': 2.27},
         'achievement_pct': 3.552, 'score': 9, 'level': '高度超預期',
         'reasons': 'Q1一季賺贏全年3倍、YoY轉虧為盈暴漲、累計超前'},
    ]
    out = render_releases_png(sample, '🆕 今日新公告 top 30 (2026-05-08)',
                              'reports/test_render.png',
                              date_str='2026-05-08',
                              first_seen_map={'5289': '2026-05-07', '4973': '2026-05-08'})
    print(f'產出: {out}')
