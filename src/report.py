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


def _format_reasons(reasons, max_items: int = 3) -> str:
    """規則版 reasons 是 list[str]，AI 版是 str。統一輸出成 '、' 連接的字串。"""
    if not reasons:
        return ''
    if isinstance(reasons, str):
        return reasons
    if isinstance(reasons, list):
        return '、'.join(str(x) for x in reasons[:max_items])
    return str(reasons)


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


def _prior_year(q_label: str) -> int:
    """'2026Q1' -> 2025"""
    return int(q_label[:4]) - 1


def _is_won_full_year(r: dict) -> bool:
    """是否確認本期累計已贏去年全年（含虧轉盈）"""
    ach = r.get('achievement_pct')
    if isinstance(ach, (int, float)) and ach >= 1.0:
        return True
    if ach == 'prior_loss' and (r.get('latest_eps') or 0) > 0:
        return True
    pf = r.get('prior_year_full')
    if pf is not None and pf < 0 and (r.get('latest_eps') or 0) > 0:
        return True
    return False


def write_summary(ws, date_str: str, stats: dict, releases: list = None, q_label: str = None):
    """摘要分頁 — 升級版含已贏全年清單 + 觀察重點 (v3 風格)"""
    ws['A1'] = f'📊 EPS 日報 {date_str}'
    ws['A1'].font = Font(name='Microsoft JhengHei', bold=True, size=16, color='1F4E78')
    ws.merge_cells('A1:F1')

    ws['A3'] = '今日新公告'; ws['B3'] = stats.get('new_count', 0)
    ws['A4'] = '評分 ≥ 8 高度超預期'; ws['B4'] = stats.get('hot_count', 0)
    ws['A5'] = '評分 6-7 值得關注'; ws['B5'] = stats.get('watch_count', 0)
    ws['A6'] = '評分 ≤ -4 衰退警示'; ws['B6'] = stats.get('warn_count', 0)
    ws['A7'] = '全市場分析筆數'; ws['B7'] = stats.get('total_count', 0)
    for r in range(3, 8):
        ws.cell(row=r, column=1).font = Font(bold=True)

    row = 9
    if releases and q_label:
        prior_year = _prior_year(q_label)
        # 已贏全年確認 list
        won = [r for r in releases
               if r.get('latest_quarter') == q_label and _is_won_full_year(r)]
        won.sort(key=lambda x: -(x.get('latest_eps') or 0))

        if won:
            ws.cell(row=row, column=1, value=f'🏆 已確認 {q_label} EPS 已贏 {prior_year} 全年 ({len(won)} 檔)')
            ws.cell(row=row, column=1).font = Font(bold=True, size=12, color='B45F06')
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            row += 1
            sub_headers = ['代號', '名稱', f'{q_label} / {prior_year}全年', '倍數', '評分', '評分理由']
            for c, h in enumerate(sub_headers, 1):
                ws.cell(row=row, column=c, value=h)
            _style_header_row(ws, row, len(sub_headers))
            row += 1
            for r in won[:10]:
                ach = r.get('achievement_pct')
                pf = r.get('prior_year_full')
                eps = r.get('latest_eps')
                if isinstance(ach, (int, float)):
                    ratio = f'{eps} / {pf} = {ach:.2f}x'
                    multi = f'{ach:.2f}x'
                else:
                    ratio = f'{eps} / {pf} (虧轉盈)' if pf is not None else f'{eps} (虧轉盈)'
                    multi = '虧轉盈'
                reasons = _format_reasons(r.get('reasons'), max_items=2)
                ws.cell(row=row, column=1, value=r.get('stock_id'))
                ws.cell(row=row, column=2, value=r.get('name', ''))
                ws.cell(row=row, column=3, value=ratio)
                ws.cell(row=row, column=4, value=multi)
                ws.cell(row=row, column=5, value=r.get('score'))
                ws.cell(row=row, column=6, value=reasons)
                for c in range(1, 7):
                    ws.cell(row=row, column=c).fill = FILL_HOT
                row += 1
            row += 1

        # 觀察重點 (bullet points)
        cur_q = [r for r in releases if r.get('latest_quarter') == q_label]
        hot8 = [r for r in cur_q if (r.get('score') or 0) >= 8]
        watch = [r for r in cur_q if 6 <= (r.get('score') or 0) <= 7]
        notable = [r for r in cur_q if 4 <= (r.get('score') or 0) <= 5]
        top3 = sorted([r for r in cur_q if r.get('latest_eps') is not None],
                      key=lambda x: -x['latest_eps'])[:3]

        ws.cell(row=row, column=1, value='📊 觀察重點')
        ws.cell(row=row, column=1).font = Font(bold=True, size=12, color='1F4E78')
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1
        bullets = []
        bullets.append(f'• 共 {len(cur_q)} 檔有 {q_label} EPS 資料；其中 +8 以上 {len(hot8)} 檔、+6~7 {len(watch)} 檔、+4~5 {len(notable)} 檔')
        if top3:
            tstr = '、'.join(f'{r.get("name") or r["stock_id"]} {r["latest_eps"]}' for r in top3)
            bullets.append(f'• {q_label} EPS 最高三名：{tstr}')
        if won:
            bullets.append(f'• 已確認 {q_label} 一季賺贏 {prior_year} 全年 {len(won)} 檔（含虧轉盈）')
        cand_count = sum(1 for r in cur_q
                         if r.get('prev_quarter_eps') is not None and r['prev_quarter_eps'] < 0
                         and (r.get('latest_eps') or 0) > 0
                         and not _is_won_full_year(r))
        if cand_count:
            bullets.append(f'• 候選名單（上季虧損、{q_label} 轉正、待全年驗證）：{cand_count} 檔（見 Sheet「{q_label}贏全年候選」）')
        for b in bullets:
            ws.cell(row=row, column=1, value=b)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
            row += 1
        row += 1

    ws.cell(row=row, column=1, value='資料來源')
    ws.cell(row=row, column=2, value='FinMind API (https://finmindtrade.com)')
    row += 1
    ws.cell(row=row, column=1, value='產出時間')
    ws.cell(row=row, column=2, value=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 26
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 50


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
            _format_reasons(r.get('reasons'), max_items=3),
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


def write_won_full_year(ws, releases: list, q_label: str):
    """已贏全年確認名單 — 累計達成率 ≥ 100% 或 虧轉盈
    對應 v3 Excel 的「Q1贏全年確認名單」分頁。
    """
    prior_year = _prior_year(q_label)
    rows = [r for r in releases
            if r.get('latest_quarter') == q_label and _is_won_full_year(r)]
    rows.sort(key=lambda x: -(x.get('latest_eps') or 0))

    title = f'🏆 已確認 {q_label} EPS > {prior_year} 全年 EPS ({len(rows)} 檔)'
    ws['A1'] = title
    ws['A1'].font = Font(name='Microsoft JhengHei', bold=True, size=14, color='B45F06')
    ws.merge_cells('A1:H1')
    ws['A1'].alignment = ALIGN_LEFT

    headers = ['代號', '名稱', f'{q_label} EPS', f'{prior_year}Q{q_label[-1]} EPS',
               f'{prior_year} 全年 EPS', '累計/全年', '倍數', '驚喜分數']
    for c, h in enumerate(headers, 1):
        ws.cell(row=3, column=c, value=h)
    _style_header_row(ws, 3, len(headers))

    for r in rows:
        ach = r.get('achievement_pct')
        pf = r.get('prior_year_full')
        if isinstance(ach, (int, float)):
            ratio_str = f'{ach*100:.1f}%'
            multi_str = f'{ach:.2f}x'
        else:
            ratio_str = '虧轉盈'
            multi_str = '虧轉盈'
        ws.append([
            r.get('stock_id'),
            r.get('name', ''),
            r.get('latest_eps'),
            r.get('yoy_eps'),
            pf,
            ratio_str,
            multi_str,
            r.get('score'),
        ])
    # 上色 (整列淡橘)
    for row_idx in range(4, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=row_idx, column=c).fill = FILL_HOT

    ws.freeze_panes = 'A4'
    _autofit(ws)


def write_won_full_year_candidates(ws, releases: list, q_label: str):
    """贏全年候選 — 上季虧損 + 本期轉正 (待全年驗證)
    對應 v3 Excel 的「Q1贏全年候選」分頁。
    """
    rows = [r for r in releases
            if r.get('latest_quarter') == q_label
            and r.get('prev_quarter_eps') is not None and r['prev_quarter_eps'] < 0
            and r.get('latest_eps') is not None and r['latest_eps'] > 0
            and not _is_won_full_year(r)]
    # 排序: QoQ Δ 由大到小
    rows.sort(key=lambda x: -((x.get('latest_eps') or 0) - (x.get('prev_quarter_eps') or 0)))

    title = f'📋 候選名單：上季虧損 + {q_label} 轉正（{len(rows)} 檔，待全年確認）'
    ws['A1'] = title
    ws['A1'].font = Font(name='Microsoft JhengHei', bold=True, size=14, color='2A6099')
    ws.merge_cells('A1:I1')
    ws['A1'].alignment = ALIGN_LEFT
    ws['A2'] = '說明：上季 (前一季) EPS 為負而本期 EPS 已轉正，若全年加總仍虧損 / 微利 則本期可能贏全年。'
    ws['A2'].font = Font(italic=True, color='666666')
    ws.merge_cells('A2:I2')

    headers = ['代號', '名稱', '驚喜分數', f'{q_label} EPS', '上季 EPS', 'QoQ Δ',
               'GM%', 'OPM%', '判斷邏輯']
    for c, h in enumerate(headers, 1):
        ws.cell(row=4, column=c, value=h)
    _style_header_row(ws, 4, len(headers))

    for r in rows:
        prev = r.get('prev_quarter_eps')
        latest = r.get('latest_eps') or 0
        qoq_delta = round(latest - (prev or 0), 2)
        ws.append([
            r.get('stock_id'),
            r.get('name', ''),
            r.get('score'),
            latest,
            prev,
            qoq_delta,
            f'{(r.get("gm_pct") or 0)*100:.1f}%' if r.get('gm_pct') is not None else '',
            f'{(r.get("opm_pct") or 0)*100:.1f}%' if r.get('opm_pct') is not None else '',
            f'{r.get("prev_quarter") or "上季"} 虧損 → {q_label} 轉正',
        ])

    # 上色: 高分淡黃
    for row_idx in range(5, ws.max_row + 1):
        sc = ws.cell(row=row_idx, column=3).value
        if isinstance(sc, (int, float)) and sc >= 6:
            for c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=c).fill = FILL_HIGHLIGHT

    ws.freeze_panes = 'A5'
    _autofit(ws)


def write_top_eps(ws, releases: list, q_label: str, top_n: int = 30):
    """高 EPS 排行 (前 N 名)
    對應 v3 Excel 的「Q1高EPS排行」分頁。
    """
    rows = [r for r in releases
            if r.get('latest_quarter') == q_label and r.get('latest_eps') is not None]
    rows.sort(key=lambda x: -x['latest_eps'])
    rows = rows[:top_n]

    title = f'💰 {q_label} EPS 高低排行（前 {len(rows)} 名）'
    ws['A1'] = title
    ws['A1'].font = Font(name='Microsoft JhengHei', bold=True, size=14, color='38761D')
    ws.merge_cells('A1:I1')
    ws['A1'].alignment = ALIGN_LEFT

    headers = ['排名', '代號', '名稱', '驚喜分數', f'{q_label} EPS', '上季 EPS', 'QoQ Δ',
               '去年全年 EPS', '累計達成率']
    for c, h in enumerate(headers, 1):
        ws.cell(row=3, column=c, value=h)
    _style_header_row(ws, 3, len(headers))

    for i, r in enumerate(rows, 1):
        prev = r.get('prev_quarter_eps')
        latest = r.get('latest_eps')
        qoq_delta = round(latest - prev, 2) if (prev is not None and latest is not None) else None
        ach = r.get('achievement_pct')
        if isinstance(ach, (int, float)):
            ach_str = f'{ach*100:.0f}%'
        elif ach == 'prior_loss':
            ach_str = '虧轉盈'
        else:
            ach_str = '—'
        ws.append([
            i,
            r.get('stock_id'),
            r.get('name', ''),
            r.get('score'),
            latest,
            prev,
            qoq_delta,
            r.get('prior_year_full'),
            ach_str,
        ])

    # 上色: 已贏全年標記
    for row_idx in range(4, ws.max_row + 1):
        ach_cell = ws.cell(row=row_idx, column=9).value
        if ach_cell == '虧轉盈' or (isinstance(ach_cell, str) and ach_cell.endswith('%')
                                       and ach_cell != '—'):
            try:
                pct = float(ach_cell.replace('%', ''))
                if pct >= 100:
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=row_idx, column=c).fill = FILL_HOT
            except ValueError:
                if ach_cell == '虧轉盈':
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=row_idx, column=c).fill = FILL_HOT

    ws.freeze_panes = 'A4'
    _autofit(ws)


def _detect_q_label(releases: list) -> str:
    """從 releases 中找最常見的 latest_quarter (即「當期」)"""
    counter = {}
    for r in releases:
        q = r.get('latest_quarter')
        if q:
            counter[q] = counter.get(q, 0) + 1
    if not counter:
        return ''
    return max(counter.items(), key=lambda kv: kv[1])[0]


def build_report(date_str: str, releases: list, monthly: list, stats: dict, out_path: str,
                 q_label: str = None):
    """產出完整報告

    分頁順序 (v3 風格):
      1. 摘要 — 已贏全年清單 + 觀察重點
      2. 今日新公告 — 完整欄位
      3. 已贏全年確認名單 — 一季賺贏全年的明星標的
      4. 贏全年候選 — 上季虧損 + 本期轉正 (待全年確認)
      5. 高 EPS 排行 — 當期 EPS 前 30
      6. 高分排行 — score ≥ 4
      7. 衰退警示 — score ≤ -4
      8. 完整 EPS — 全市場
      9. 月營收
    """
    if not q_label:
        q_label = _detect_q_label(releases)

    wb = Workbook()
    # 摘要 (含已贏全年清單 + 觀察重點)
    ws_sum = wb.active
    ws_sum.title = '摘要'
    write_summary(ws_sum, date_str, stats, releases=releases, q_label=q_label)

    # 今日新公告 (依分數排序)
    new_only = sorted([r for r in releases if r.get('is_new')],
                      key=lambda x: -(x.get('score') or -99))
    if new_only:
        ws_new = wb.create_sheet('今日新公告')
        write_releases(ws_new, new_only, '今日新公告')

    # v3 風格三個 sheet — 都依 q_label 篩選
    if q_label:
        won = [r for r in releases
               if r.get('latest_quarter') == q_label and _is_won_full_year(r)]
        if won:
            ws_won = wb.create_sheet(f'{q_label}贏全年確認')
            write_won_full_year(ws_won, releases, q_label)

        cand = [r for r in releases
                if r.get('latest_quarter') == q_label
                and r.get('prev_quarter_eps') is not None and r['prev_quarter_eps'] < 0
                and (r.get('latest_eps') or 0) > 0
                and not _is_won_full_year(r)]
        if cand:
            ws_cand = wb.create_sheet(f'{q_label}贏全年候選')
            write_won_full_year_candidates(ws_cand, releases, q_label)

        cur_q = [r for r in releases if r.get('latest_quarter') == q_label
                 and r.get('latest_eps') is not None]
        if cur_q:
            ws_top = wb.create_sheet(f'{q_label}高EPS排行')
            write_top_eps(ws_top, releases, q_label, top_n=30)

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
