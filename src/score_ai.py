"""
AI 驚喜度評分 (-9 ~ +9) — Anthropic Claude API

依使用者定義：
  ① 驚喜度 — 單季 EPS vs 上季：大幅成長 = 驚喜
  ② 延續性 — 獲利趨勢能否持續
  ③ 獲利品質 — 本業驅動 > 業外收益

級距:
  +8~9 高度超預期 | +6~7 值得關注 | +4~5 有亮點待觀察
  +1~3 符合預期 | 0 無特別資訊 | 負分 衰退警示

成本優化:
- 用 prompt caching 把 system + 評分標準 + few-shot 範例固定快取 (~5 分鐘 TTL)
- 每筆 input 只送變動的股票分析資料 (~150 tokens)
- 用 messages.parse() + Pydantic schema 確保輸出結構
"""
import os
import json
from typing import Optional
from pydantic import BaseModel, Field

import anthropic


SYSTEM_PROMPT = """你是台股財報驚喜度評分專家。

你的任務：依據單檔股票最新季財報資料，給出「驚喜度評分」-9 到 +9 整數，並用一句話說明理由。

## 評估面向（三大維度）

### ① 驚喜度 (Surprise)
- 單季 EPS vs 上季：大幅成長 = 驚喜，大幅衰退 = 反向驚喜
- 單季 EPS vs 去年同季：YoY 成長率反映基期變化
- 轉虧為盈 / 由盈轉虧 是極強的訊號

### ② 延續性 (Sustainability)
- 累計達成率 (今年累計 EPS / 去年全年 EPS)：> 100% = 已賺贏全年
- Q1 占去年全年比例 vs 同期 baseline (常態約 25%)：超過 50% = 大幅超前
- 連續成長 vs 一次性跳升

### ③ 獲利品質 (Quality)
- 業外比例 (業外損益 / 稅後淨利)：< 30% = 本業驅動，> 50% = 業外灌水
- 毛利率 (GM%)、營業利益率 (OPM%)：高且穩 = 體質好
- 本業虧損 (OPM < 0) = 即使 EPS 為正也應扣分

## 評分級距

| 分數 | 級別 | 說明 |
|---|---|---|
| +8~+9 | 高度超預期 | 三面向都很強：QoQ/YoY 暴衝、累計超前、本業驅動 |
| +6~+7 | 值得關注 | 兩面向強烈正向、第三面向至少不負分 |
| +4~+5 | 有亮點待觀察 | 一面向明顯正向，其他面向中性 |
| +1~+3 | 符合預期 | 小幅成長或穩定，無特別亮點 |
| 0 | 無特別資訊 | 平盤、無變化 |
| -1~-3 | 衰退警示 | 一面向負向 |
| -4~-6 | 明顯衰退 | 兩面向負向 |
| -7~-9 | 嚴重衰退 | 三面向都負，或由盈轉鉅虧 |

## 重要規則

1. **EPS 正負本身不重要**，重要的是「相對」變化 (QoQ, YoY) 和品質指標
2. **單季高 EPS 但業外貢獻 > 50%** → 即使分數正向也要降 1-2 分標記
3. **Q1 一季就賺贏全年** → 自動 +6 起跳
4. **轉虧為盈** → 自動 +5 起跳
5. **理由請用中文，一句話 (15 字以內)**，例如「QoQ +671%、本業驅動」
6. **嚴格按照 JSON schema 輸出**，不要加 markdown code block。
"""

# Few-shot 範例 (cache 友善 — 不會變)
FEW_SHOT_EXAMPLES = """
## 評分範例 (學習用)

範例 1: 5386 青雲 2026 Q1
- 最新 EPS: 43.05 (2025 Q1 為 1.27, 2025 Q4 為 5.58)
- QoQ: +671%, YoY: +3290%
- 累計達成率: 490% (Q1 一季就賺贏 2025 全年 8.78 約 5 倍)
- GM 27%, OPM 23%, 業外比 -2% (本業強勁)
→ {"score": 9, "level": "高度超預期", "reasons": "QoQ 爆發+671%、Q1 一季賺贏全年 5 倍、本業驅動"}

範例 2: 6640 均華 2026 Q1
- 最新 EPS: 5.19 (2025 Q1 為 1.64, 2025 Q4 為 3.35)
- QoQ: +55%, YoY: +217%
- 累計達成率: 41% (5.19 / 12.75)
- GM 45%, OPM 22%, 業外比 4%
→ {"score": 6, "level": "值得關注", "reasons": "YoY 大增+217%、本業強勁但達成率正常"}

範例 3: 6224 聚鼎 2026 Q1
- 最新 EPS: 0.64 (2025 Q1 為 0.43, 2025 Q4 為 -0.30)
- QoQ: 轉虧為盈, YoY: +49%
- 累計達成率: 97% (Q1 一季差點贏全年)
- GM 32%, OPM 6%, 業外比 37%
→ {"score": 7, "level": "值得關注", "reasons": "Q4 轉虧為盈、累計近全年但業外偏高"}

範例 4: 8467 波力-KY 2026 Q1
- 最新 EPS: 0.83 (2025 Q1 為 2.60, 2025 Q4 為 0.61)
- QoQ: +36%, YoY: -68%
- 累計達成率: 8% (進度落後)
- GM 37%, OPM 17%, 業外比 -85% (鉅額業外損失)
→ {"score": -3, "level": "衰退警示", "reasons": "YoY 大跌 -68%、業外鉅虧侵蝕本業"}
"""


class ScoreOutput(BaseModel):
    """AI 評分輸出 schema"""
    score: int = Field(description="驚喜度評分 -9 ~ +9 整數", ge=-9, le=9)
    level: str = Field(description="級別：高度超預期/值得關注/有亮點待觀察/符合預期/無特別資訊/衰退警示/明顯衰退/嚴重衰退")
    reasons: str = Field(description="評分理由，一句話 15 字內")


def _format_input(analysis: dict) -> str:
    """把分析資料 format 成 prompt input"""
    sid = analysis.get('stock_id', '')
    name = analysis.get('name', '')
    q = analysis.get('latest_quarter', '')
    eps = analysis.get('latest_eps')
    yoy_eps = analysis.get('yoy_eps')
    yoy = analysis.get('yoy') or {}
    qoq = analysis.get('qoq') or {}
    accum = analysis.get('accumulated') or {}
    prior_full = analysis.get('prior_year_full')
    ach = analysis.get('achievement_pct')
    gm = analysis.get('gm_pct')
    opm = analysis.get('opm_pct')
    nonop = analysis.get('nonop_pct')

    yoy_str = f'{yoy.get("pct", 0)*100:+.0f}%' if yoy.get('pct') is not None else 'N/A'
    qoq_str = f'{qoq.get("pct", 0)*100:+.0f}%' if qoq.get('pct') is not None else 'N/A'
    ach_str = (f'{ach*100:.0f}%' if isinstance(ach, (int, float))
               else ('去年虧損' if ach == 'prior_loss' else 'N/A'))
    gm_str = f'{gm*100:.1f}%' if gm is not None else 'N/A'
    opm_str = f'{opm*100:.1f}%' if opm is not None else 'N/A'
    nonop_str = f'{nonop*100:+.0f}%' if nonop is not None else 'N/A'

    return (f"代號: {sid} {name}\n"
            f"最新季: {q}, EPS: {eps}\n"
            f"去年同季 EPS: {yoy_eps}, YoY: {yoy_str}\n"
            f"QoQ: {qoq_str}\n"
            f"累計 EPS: {accum.get('value')}, 去年全年 EPS: {prior_full}\n"
            f"累計達成率: {ach_str}\n"
            f"GM: {gm_str}, OPM: {opm_str}, 業外比: {nonop_str}")


def score_one(client: anthropic.Anthropic, analysis: dict, model: str) -> dict:
    """單檔評分。
    回傳: {score, level, reasons, label, ai: True}
    使用 prompt caching:
      - system: 評分標準 + few-shot (cached, ~5 分鐘 TTL)
      - user: 變動的股票資料
    """
    if not analysis.get('has_data'):
        return {'score': None, 'level': '無資料', 'reasons': '無 EPS 資料', 'label': '—', 'ai': False}

    user_input = _format_input(analysis)

    response = client.messages.parse(
        model=model,
        max_tokens=200,
        system=[
            {
                'type': 'text',
                'text': SYSTEM_PROMPT + FEW_SHOT_EXAMPLES,
                'cache_control': {'type': 'ephemeral'},  # 5 min TTL
            }
        ],
        messages=[{'role': 'user', 'content': f'請評分以下股票：\n\n{user_input}'}],
        output_format=ScoreOutput,
    )

    parsed: ScoreOutput = response.parsed_output
    score = parsed.score
    level = parsed.level
    reasons = parsed.reasons

    # 級別 + emoji 對照
    if score >= 8:
        label = f'🔥 +{score}'
    elif score >= 6:
        label = f'⭐ +{score}'
    elif score >= 4:
        label = f'✨ +{score}'
    elif score >= 1:
        label = f'➕ +{score}'
    elif score == 0:
        label = '➖ 0'
    elif score >= -3:
        label = f'⚠️ {score}'
    else:
        label = f'🔻 {score}'

    return {
        'score': score,
        'level': level,
        'reasons': reasons,
        'label': label,
        'ai': True,
        '_usage': {
            'cache_read': response.usage.cache_read_input_tokens,
            'cache_create': response.usage.cache_creation_input_tokens,
            'input': response.usage.input_tokens,
            'output': response.usage.output_tokens,
        },
    }


def score_batch(analyses: list, api_key: Optional[str] = None,
                model: Optional[str] = None, progress: bool = True) -> list:
    """批次評分。傳回 list of {score, level, reasons, label, ai, _usage}."""
    client = anthropic.Anthropic(api_key=api_key or os.environ['ANTHROPIC_API_KEY'])
    model = model or os.environ.get('SCORE_MODEL', 'claude-haiku-4-5')
    results = []
    total = len(analyses)
    cache_reads = 0
    for i, a in enumerate(analyses):
        if progress and i % 10 == 0:
            print(f'  [AI {i}/{total}] {a.get("stock_id")} {a.get("name", "")}...', flush=True)
        try:
            r = score_one(client, a, model=model)
            cache_reads += (r.get('_usage') or {}).get('cache_read', 0)
        except anthropic.RateLimitError:
            import time
            time.sleep(5)
            r = score_one(client, a, model=model)
        except Exception as e:
            print(f'  ⚠️ {a.get("stock_id")} 評分失敗: {e}', flush=True)
            r = {'score': None, 'level': '評分失敗', 'reasons': str(e)[:50], 'label': '⚠ ?', 'ai': False}
        results.append(r)
    if progress:
        print(f'  AI 評分完成: {total} 筆, 快取讀取 {cache_reads} tokens')
    return results


if __name__ == '__main__':
    # 自測
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    from pathlib import Path
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / '.env', override=True)

    test_cases = [
        {
            'stock_id': '5386', 'name': '青雲', 'has_data': True,
            'latest_quarter': '2026Q1', 'latest_eps': 43.05, 'yoy_eps': 1.27,
            'yoy': {'pct': 32.9}, 'qoq': {'pct': 6.71},
            'accumulated': {'value': 43.05, 'quarters_count': 1},
            'prior_year_full': 8.78, 'achievement_pct': 4.9,
            'gm_pct': 0.27, 'opm_pct': 0.23, 'nonop_pct': -0.02,
        },
        {
            'stock_id': '8467', 'name': '波力-KY', 'has_data': True,
            'latest_quarter': '2026Q1', 'latest_eps': 0.83, 'yoy_eps': 2.60,
            'yoy': {'pct': -0.68}, 'qoq': {'pct': 0.36},
            'accumulated': {'value': 0.83, 'quarters_count': 1},
            'prior_year_full': 10.04, 'achievement_pct': 0.083,
            'gm_pct': 0.37, 'opm_pct': 0.17, 'nonop_pct': -0.85,
        },
    ]
    results = score_batch(test_cases)
    for case, r in zip(test_cases, results):
        print(f'\n{case["stock_id"]} {case["name"]}:')
        print(f'  分數: {r.get("score")} ({r.get("level")})')
        print(f'  理由: {r.get("reasons")}')
        print(f'  使用量: {r.get("_usage")}')
