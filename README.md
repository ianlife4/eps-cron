# EPS Cron — 台股財報每日監控

每天自動爬取上市櫃 EPS / 月營收，計算 YoY、累計達成率，產出 Excel 報告並推到 Telegram。

## 架構

```
資料源: FinMind API (主) + MOPS 公開資訊觀測站 (備)
排程:   GitHub Actions cron (每天 18:00 + 21:30 UTC+8)
分析:   YoY、累計達成率、規則版驚喜度評分 (-9 ~ +9)
派送:   Telegram Bot (摘要 + Excel 附件) + Cloudflare Worker /eps 頁
```

## 目錄結構

```
src/
  fetch_eps.py        # FinMind EPS 爬蟲
  fetch_revenue.py    # 月營收爬蟲
  fetch_self.py       # 自結 EPS 爬蟲 (MOPS)
  compare.py          # YoY、累計達成率
  score.py            # 規則版驚喜度評分 (AI 版作備換)
  report.py           # Excel 報告產生器
  notify.py           # Telegram 推播
  main.py             # 主控
data/
  snapshots/          # 每日 JSON snapshot (idempotent skip 用)
  historical/         # 歷史 EPS / 營收資料庫
reports/              # 每日 Excel 輸出
.github/workflows/
  cron.yml            # 排程設定
```

## 本機開發

```bash
pip install -r requirements.txt
cp .env.template .env     # 填入 TG_BOT_TOKEN 等
python src/main.py        # 跑一次
```

## 部署

GitHub Secrets 需要：
- `TG_BOT_TOKEN`
- `TG_CHAT_ID`
- `ANTHROPIC_API_KEY` (可選，啟用 AI 評分)
