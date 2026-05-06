# 部署指南

## GitHub 部署 (推薦)

### Step 1: 建 Repo
```bash
cd "C:\Users\J.Chun\Desktop\eps-cron"
git init
git add -A
git commit -m "Initial commit: EPS Daily Monitor"
gh repo create eps-cron --private --source=. --push
```

### Step 2: 設定 Secrets
```bash
# Telegram Bot
gh secret set TG_BOT_TOKEN --body "8601835191:AAEdC95Z..."
gh secret set TG_CHAT_ID --body "2103179843"

# Anthropic API
gh secret set ANTHROPIC_API_KEY --body "sk-ant-api03-..."
```

或進入 GitHub → Settings → Secrets and variables → Actions 手動加。

### Step 3: 啟用 Workflow
GitHub Actions 預設啟用。第一次跑可以手動觸發：
```bash
gh workflow run daily.yml -f force_all=true -f max_ai=50
```

或進入 Actions tab → "EPS Daily Cron" → Run workflow。

---

## 排程

預設兩次：
- **18:00 TW (10:00 UTC)** 週一至五 — 證交所收盤後
- **21:30 TW (13:30 UTC)** 週一至五 — 櫃買晚間公告

## 停用排程
進入 Actions tab → "EPS Daily Cron" → "..." → Disable workflow

## 改時段
編輯 `.github/workflows/daily.yml` 的 `cron:` 行 (UTC 時間)。
