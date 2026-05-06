"""
Excel 報告產生器
產出多分頁 Excel：
  摘要 / 今日新公告 / 完整 EPS / 月營收 / 排行榜
"""
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 樣式
FONT_HEADER = Font(name='Microsoft JhengHei', bold=True, color='FFFFFF', size=11)
FILL_HEADER = PatternFill('solid', start_color='1F4E78')
FILL_HIGHLIGHT = PatternFill('solid', start_color='FFF2CC')
FILL_HOT = PatternFill('solid', start_color='FCE4D6')
FILL_COLD = PatternFill('solid', start_color='D9E1F2')
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
THIN = Side(border_style='thin', color='BFBFBF')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def _style_header_row(ws, row: int, cols: int):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER


def _autofit(ws, min_width: int = 8, max_width: int = 30):
    for col_idx, col in enumerate(ws.columns, 1):
        max_len = min_width
        for cell in col:
            try:
                v = str(cell.value) if cell.value is not None else ''
                # 中文佔 2 寬度
                w = sum(2 if ord(ch) > 127 else 1 for ch in v)
                max_len = max(max_len, w)
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, max_width)


def write_summary(ws, date_str: str, stats: dict):
    """摘要分頁"""
    ws['A1'] = f'📊 EPS 日報 {date_str}'
    ws['A1'].font = Font(name='Microsoft JhengHei', bold=True, size=16, color='1F4E78')
    ws.merge_cells('A1:F1')

    ws['A3'] = '今日新公告'; ws['B3'] = stats.get('new_count', 0)
    ws['A4'] = '評分 ≥ 8 高度超預期'; ws['B4'] = stats.get('hot_count', 0)
    ws['A5'] = '評分 6-7 值得關注'; ws['B5'] = stats.get('watch_count', 0)
    ws['A6'] = '評分 ≤ -4 衰退警示'; ws['B6'] = stats.get('warn_count', 0)
    ws['A7'] = '全市場分析筆數'; ws['B7'] = stats.get('total_count', 0)

    ws['A9'] = '資料來源'; ws['B9'] = 'FinMind API (https://finmindtrade.com)'
    ws['A10'] = '產出時間'; ws['B10'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    for r in range(3, 11):
        ws.cell(row=r, column=1).font = Font(bold=True)
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 50


def write_releases(ws, releases: list, title: str = '今日新公告'):
    """新公告分頁（包含 stock_id, name, market, latest_q, eps, YoY, 達成率, score, level, reasons）"""
    headers = ['代號', '名稱', '市場', '最新季', 'EPS', '去年同季 EPS', 'YoY %', 'YoY Δ',
               '累計 EPS', '去年全年', '達成率 %', 'QoQ %',
               '驚喜分數', '級別', '評分理由', 'GM%', 'OPM%', '業外%']
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    for r in releases:
        yoy = r.get('yoy') or {}
        qoq = r.get('qoq') or {}
        accum = r.get('accumulated') or {}
        ach = r.get('achievement_pct')
        ach_str = (f'{ach*100:.1f}%' if isinstance(ach, (int, float)) else
                   ('去年虧損' if ach == 'prior_loss' else '—'))
        ws.append([
            r.get('stock_id'),
            r.get('name', ''),
            r.get('market', ''),
            r.get('latest_quarter', ''),
            r.get('latest_eps'),
            r.get('yoy_eps'),
            f'{yoy.get("pct", 0)*100:+.1f}%' if yoy.get('pct') is not None else '—',
            yoy.get('delta'),
            accum.get('value'),
            r.get('prior_year_full'),
            ach_str,
            f'{qoq.get("pct", 0)*100:+.1f}%' if qoq.get('pct') is not None else '—',
            r.get('score'),
            r.get('level', ''),
            '、'.join((r.get('reasons') or [])[:3]),
            f'{(r.get("gm_pct") or 0)*100:.1f}%' if r.get('gm_pct') is not None else '',
            f'{(r.get("opm_pct") or 0)*100:.1f}%' if r.get('opm_pct') is not None else '',
            f'{(r.get("nonop_pct") or 0)*100:.1f}%' if r.get('nonop_pct') is not None else '',
        ])

    # 上色
    for row_idx in range(2, ws.max_row + 1):
        score = ws.cell(row=row_idx, column=13).value
        fill = None
        if isinstance(score, (int, float)):
            if score >= 8: fill = FILL_HOT
            elif score >= 4: fill = FILL_HIGHLIGHT
            elif score <= -4: fill = FILL_COLD
        if fill:
            for c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=c).fill = fill

    ws.freeze_panes = 'A2'
    _autofit(ws)


def write_revenue(ws, monthly_data: list, title: str = '月營收'):
    """月營收分頁
    monthly_data: list of dicts:
      {stock_id, name, market, ym (YYYY-MM), revenue, yoy, mom, accum, accum_yoy}
    """
    headers = ['代號', '名稱', '市場', '月份', '當月營收(千元)', 'YoY %', 'MoM %',
               '累計營收(千元)', '累計 YoY %']
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    for r in monthly_data:
        rev = r.get('revenue')
        rev_k = round(rev / 1000) if rev else None
        accum = r.get('accumulated')
        accum_k = round(accum / 1000) if accum else None
        yoy = r.get('yoy')
        yoy_str = f'{yoy:+.1f}%' if yoy is not None else '—'
        mom = r.get('mom')
        mom_str = f'{mom:+.1f}%' if mom is not None else '—'
        a_yoy = r.get('accum_yoy')
        a_yoy_str = f'{a_yoy:+.1f}%' if a_yoy is not None else '—'
        ws.append([
            r.get('stock_id'),
            r.get('name', ''),
            r.get('market', ''),
            r.get('ym', ''),
            rev_k,
            yoy_str,
            mom_str,
            accum_k,
            a_yoy_str,
        ])

    # 上色: YoY ≥ +50%
    for row_idx in range(2, ws.max_row + 1):
        cell_yoy = ws.cell(row=row_idx, column=6).value
        try:
            yoy_v = float(str(cell_yoy).replace('%', '').replace('+', ''))
            if yoy_v >= 50:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=c).fill = FILL_HOT
            elif yoy_v <= -30:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=row_idx, column=c).fill = FILL_COLD
        except (ValueError, TypeError):
            pass

    ws.freeze_panes = 'A2'
    _autofit(ws)


def build_report(date_str: str, releases: list, monthly: list, stats: dict, out_path: str):
    """產出完整報告"""
    wb = Workbook()
    # 摘要
    ws_sum = wb.active
    ws_sum.title = '摘要'
    write_summary(ws_sum, date_str, stats)

    # 今日新公告 (依分數排序)
    new_only = sorted([r for r in releases if r.get('is_new')],
                      key=lambda x: -(x.get('score') or -99))
    if new_only:
        ws_new = wb.create_sheet('今日新公告')
        write_releases(ws_new, new_only, '今日新公告')

    # 高分排行 (score ≥ 4)
    hot = sorted([r for r in releases if (r.get('score') or 0) >= 4],
                 key=lambda x: -(x.get('score') or 0))
    if hot:
        ws_hot = wb.create_sheet('高分排行')
        write_releases(ws_hot, hot, '高分排行')

    # 衰退警示 (score ≤ -4)
    cold = sorted([r for r in releases if (r.get('score') or 0) <= -4],
                  key=lambda x: x.get('score') or 0)
    if cold:
        ws_cold = wb.create_sheet('衰退警示')
        write_releases(ws_cold, cold, '衰退警示')

    # 完整 EPS
    full = sorted(releases, key=lambda x: -(x.get('score') or -99))
    ws_full = wb.create_sheet('完整 EPS')
    write_releases(ws_full, full, '完整 EPS')

    # 月營收
    if monthly:
        ws_rev = wb.create_sheet('月營收')
        write_revenue(ws_rev, monthly)

    # 移除無印表機設定避免開檔卡
    from openpyxl.worksheet.page import PageMargins, PrintOptions, PrintPageSetup
    for sname in wb.sheetnames:
        ws = wb[sname]
        new_setup = PrintPageSetup(worksheet=ws)
        new_setup.paperSize = 9
        new_setup.orientation = 'landscape'
        ws.page_setup = new_setup
        ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
        ws.print_options = PrintOptions()
        ws.print_area = None

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


if __name__ == '__main__':
    # 自測
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    sample_releases = [
        {
            'stock_id': '5386', 'name': '青雲', 'market': '上市',
            'is_new': True,
            'latest_quarter': '2026Q1', 'latest_eps': 43.05, 'yoy_eps': 1.27,
            'yoy': {'pct': 32.9, 'delta': 41.78},
            'qoq': {'pct': 6.71, 'delta': 37.47},
            'accumulated': {'value': 43.05, 'quarters_count': 1},
            'prior_year_full': 8.78, 'achievement_pct': 4.9,
            'gm_pct': 0.27, 'opm_pct': 0.23, 'nonop_pct': -0.02,
            'score': 9, 'level': '高度超預期', 'label': '🔥 +9',
            'reasons': ['QoQ 爆發 +671%', 'YoY 爆發 +3290%', '已賺贏去年全年 (490%)'],
        },
    ]
    out = build_report('2026-05-06', sample_releases, [],
                       {'new_count': 1, 'hot_count': 1, 'total_count': 1},
                       'reports/test_report.xlsx')
    print(f'產出: {out}')
