"""
驚喜度評分 (-9 ~ +9) — 規則版
仿用使用者圖示定義：
  ① 驚喜度 — 單季 EPS vs 上季：大幅成長 = 驚喜
  ② 延續性 — 獲利趨勢能否持續
  ③ 獲利品質 — 本業驅動 > 業外收益

未來可換 AI 版（同 interface），輸出格式一致。
"""
from typing import Optional


def score_rule_based(analysis: dict) -> dict:
    """
    輸入: compare.analyze_one() 的結果
    輸出: {score, level, reasons, label}
      score: -9 ~ +9 整數
      level: 高度超預期 / 值得關注 / 有亮點待觀察 / 符合預期 / 無特別資訊 / 衰退警示
      reasons: list of strings
      label: emoji + 級別 (給 TG / Excel 用)
    """
    if not analysis.get('has_data'):
        return {'score': None, 'level': '無資料', 'reasons': [], 'label': '—'}

    score = 0
    reasons = []

    yoy = analysis.get('yoy') or {}
    yoy_pct = yoy.get('pct')
    qoq = analysis.get('qoq') or {}
    qoq_pct = qoq.get('pct')
    achievement = analysis.get('achievement_pct')
    nonop_pct = analysis.get('nonop_pct')
    gm = analysis.get('gm_pct')
    opm = analysis.get('opm_pct')
    latest = analysis.get('latest_eps') or 0
    yoy_eps = analysis.get('yoy_eps')
    prior_full = analysis.get('prior_year_full')

    # ① 驚喜度 (QoQ + YoY)
    if qoq_pct is not None:
        if qoq_pct >= 2.0:           # QoQ +200% 以上
            score += 3; reasons.append(f'QoQ 爆發 +{qoq_pct*100:.0f}%')
        elif qoq_pct >= 1.0:         # QoQ +100%
            score += 2; reasons.append(f'QoQ 大增 +{qoq_pct*100:.0f}%')
        elif qoq_pct >= 0.3:
            score += 1; reasons.append(f'QoQ +{qoq_pct*100:.0f}%')
        elif qoq_pct <= -0.3:
            score -= 1; reasons.append(f'QoQ {qoq_pct*100:.0f}%')

    if yoy_pct is not None:
        if yoy_pct >= 2.0:
            score += 3; reasons.append(f'YoY 爆發 +{yoy_pct*100:.0f}%')
        elif yoy_pct >= 1.0:
            score += 2; reasons.append(f'YoY 大增 +{yoy_pct*100:.0f}%')
        elif yoy_pct >= 0.3:
            score += 1; reasons.append(f'YoY +{yoy_pct*100:.0f}%')
        elif yoy_pct <= -0.3:
            score -= 2; reasons.append(f'YoY {yoy_pct*100:.0f}%')
        elif yoy_pct <= -0.1:
            score -= 1; reasons.append(f'YoY {yoy_pct*100:.0f}%')

    # 轉虧為盈
    if yoy_eps is not None and yoy_eps < 0 and latest > 0:
        score += 2; reasons.append('轉虧為盈')
    elif yoy_eps is not None and yoy_eps > 0 and latest < 0:
        score -= 3; reasons.append('由盈轉虧')

    # ② 延續性 — 累計達成率
    if isinstance(achievement, (int, float)):
        # 例如 Q1 達 50% → 同期 baseline ~ 25%，等於超出 25pp
        if achievement >= 1.0:
            score += 3; reasons.append(f'已賺贏去年全年 ({achievement*100:.0f}%)')
        elif achievement >= 0.5 and analysis.get('accumulated', {}).get('quarters_count') == 1:
            # Q1 一季就達 50% 等於 baseline 倍 = 大幅超前
            score += 2; reasons.append(f'Q1 達 {achievement*100:.0f}% 大超前')
        elif achievement >= 0.4 and analysis.get('accumulated', {}).get('quarters_count') <= 2:
            score += 1; reasons.append(f'累計達 {achievement*100:.0f}%')
    elif achievement == 'prior_loss' and latest > 0:
        score += 2; reasons.append('去年虧損 / 今年轉正')

    # ③ 獲利品質 — 業外比例
    if nonop_pct is not None:
        if abs(nonop_pct) > 0.5:
            score -= 2; reasons.append(f'業外比 {nonop_pct*100:.0f}% 過高')
        elif abs(nonop_pct) > 0.3:
            score -= 1; reasons.append(f'業外比 {nonop_pct*100:.0f}% 偏高')

    # 毛利率改善（與去年同期比，需要拿到去年同期的 GM 才能比）
    # 暫時用絕對門檻
    if opm is not None:
        if opm >= 0.30:
            score += 1; reasons.append(f'OPM {opm*100:.0f}% 強')
        elif opm < 0:
            score -= 1; reasons.append(f'本業虧損 OPM {opm*100:.0f}%')

    # clamp
    score = max(-9, min(9, score))

    # 級別判定 (依使用者圖示)
    if score >= 8:
        level = '高度超預期'; label = '🔥 +' + str(score)
    elif score >= 6:
        level = '值得關注'; label = '⭐ +' + str(score)
    elif score >= 4:
        level = '有亮點待觀察'; label = '✨ +' + str(score)
    elif score >= 1:
        level = '符合預期'; label = '➕ +' + str(score)
    elif score == 0:
        level = '無特別資訊'; label = '➖ 0'
    elif score >= -3:
        level = '衰退警示'; label = '⚠️ ' + str(score)
    else:
        level = '嚴重衰退'; label = '🔻 ' + str(score)

    return {
        'score': score,
        'level': level,
        'reasons': reasons,
        'label': label,
    }


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    # 自測：模擬 5386 青雲 2026 Q1 (43.05) vs 2025 Q1 (1.27)
    test_analysis = {
        'has_data': True,
        'latest_eps': 43.05,
        'yoy_eps': 1.27,
        'yoy': {'pct': 32.9, 'delta': 41.78},
        'qoq': {'pct': 6.71, 'delta': 37.47},  # vs Q4 2025 = 5.58
        'achievement_pct': 4.9,  # 43.05 / 8.78 = 4.9
        'accumulated': {'quarters_count': 1},
        'prior_year_full': 8.78,
        'gm_pct': 0.27,
        'opm_pct': 0.23,
        'nonop_pct': -0.02,
    }
    r = score_rule_based(test_analysis)
    import json
    print('5386 青雲 模擬:')
    print(json.dumps(r, ensure_ascii=False, indent=2))
