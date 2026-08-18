# -*- coding: utf-8 -*-
"""法說/財報 中午提醒（加料版）— @Ianke_eps_daily_bot
- 事件: aurora research-map/events.json (外資研究地圖每日 19:30 更新)
- 財報過沒過: TWSE/TPEX OpenAPI 季度損益表 — 已申報本季 = ✅, 未申報 = 🔴 置頂
- 每檔附: 覆蓋外資 ｜上次報告 + 判讀（已反應/尚無反應→可能補報告）
- 今天+明天都沒場次 → 靜默不推
"""
import json, os, sys, urllib.request, urllib.parse, datetime

EVENTS_URL = "https://raw.githubusercontent.com/ianlife4/aurora/main/research-map/events.json"
TOKEN = os.environ["TG_BOT_TOKEN"]
CHAT = os.environ["TG_CHAT_ID"]

SHORT = {"Morgan Stanley": "MS", "Goldman Sachs": "GS", "J.P. Morgan": "JPM",
         "Citi": "Citi", "UBS": "UBS", "Daiwa": "Daiwa", "Nomura": "NMR",
         "BofA": "BofA", "Macquarie": "MQ", "CLSA": "CLSA", "HSBC": "HSBC",
         "Aletheia": "Aletheia", "Bernstein": "Bern", "Jefferies": "Jef",
         "DBS": "DBS", "廣發證券": "GF"}
WD = ["一", "二", "三", "四", "五", "六", "日"]

def fetch_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    return json.load(urllib.request.urlopen(req, timeout=60))

# ---------- 財報申報名單 (一般業+金融各業別, 上市+上櫃) ----------
def filed_map():
    """code -> (年度, 季別) 已申報的最新季"""
    out = {}
    suffixes = ["ci", "basi", "bd", "fh", "ins", "mim"]
    for sfx in suffixes:
        for base, ck, yk, sk in [
            (f"https://openapi.twse.com.tw/v1/opendata/t187ap06_L_{sfx}", "公司代號", "年度", "季別"),
            (f"https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_{sfx}", "SecuritiesCompanyCode", "Year", "Season"),
        ]:
            try:
                for r in fetch_json(base):
                    c = str(r.get(ck) or "").strip()
                    try:
                        yq = (int(r.get(yk)), int(r.get(sk)))
                    except (TypeError, ValueError):
                        continue
                    if c and yq > out.get(c, (0, 0)):
                        out[c] = yq
            except Exception as ex:
                print(f"warn: {base.split('/')[-1]} {ex}", file=sys.stderr)
    return out

def need_quarter(d):
    """這場法說對應要看的財報季 (ROC年, 季)"""
    y, m = d.year - 1911, d.month
    if m <= 3:  return (y - 1, 4)   # 1-3月 → 去年年報
    if m <= 6:  return (y, 1)
    if m <= 9:  return (y, 2)
    return (y, 3)

# ---------- 主流程 ----------
now = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
today = now.date()
tomorrow = today + datetime.timedelta(days=1)

data = fetch_json(EVENTS_URL)
evs = data.get("officialEvents", [])

def send(text):
    payload = urllib.parse.urlencode({"chat_id": CHAT, "text": text}).encode()
    r = urllib.request.urlopen(
        urllib.request.Request(f"https://api.telegram.org/bot{TOKEN}/sendMessage", data=payload), timeout=30)
    resp = json.load(r)
    print("sent:", resp.get("ok"), "msg_id:", resp.get("result", {}).get("message_id"))

# ── 看門狗:資料超過 2 天沒更新 → 告警(獨立於本機,斷更一定有人喊) ──
stale_warn = None
gen = (data.get("generated") or "")[:10]
try:
    age = (today - datetime.date.fromisoformat(gen)).days
    if age >= 2:
        stale_warn = (f"⚠ 場次資料已 {age} 天未更新（最後 {data.get('generated')}）\n"
                      f"本機 19:30 排程可能斷了：開電腦檢查 Desktop\\外資報告分析\\update.log，"
                      f"或手動跑 daily_update.bat")
except ValueError:
    stale_warn = f"⚠ 場次資料時間戳異常: {data.get('generated')!r}"

td = sorted([e for e in evs if e.get("date") == today.isoformat()], key=lambda e: e.get("time") or "99")
tm = sorted([e for e in evs if e.get("date") == tomorrow.isoformat()], key=lambda e: e.get("time") or "99")
if not td and not tm:
    if stale_warn:
        send(stale_warn)   # 沒場次的日子也要讓斷更浮出來
    else:
        print("no events today/tomorrow, skip")
    sys.exit(0)

filed = filed_map()
print(f"filed companies: {len(filed)}")

def is_filed(e, d):
    return filed.get(e["code"], (0, 0)) >= need_quarter(d)

def loc_tag(e):
    loc = (e.get("loc") or "").strip()
    online = any(k in loc for k in ("線上", "Webex", "WEBEX", "webex", "Teams", "TEAMS", "電話", "視訊"))
    return "線上" if online else (loc[:10] if loc else "")

def brokers(e):
    return "·".join(SHORT.get(b, b) for b in (e.get("reactors") or [])[:5])

def md(iso):
    return f"{int(iso[5:7])}/{int(iso[8:10])}" if iso else "?"

def line1(e, mark=""):
    t = (e.get("time") or "").strip()
    lt = loc_tag(e)
    return (f"{t} " if t else "") + f"{e['code']} {e['name']}" + (f"（{lt}）" if lt else "") + mark

def line2(e, filed_flag):
    lr = f"上次報告 {md(e.get('lr_d'))} {SHORT.get(e.get('lr_b'), e.get('lr_b') or '')}" if e.get("lr_d") else "尚無報告"
    if filed_flag:
        verdict = "（財報已反應）" if e.get("reacted") else "（尚無反應→這場可能補）"
    else:
        verdict = ""
    return f"　　{brokers(e) or '—'} ｜{lr}{verdict}"

red = [e for e in td if not is_filed(e, today)]
grn = [e for e in td if is_filed(e, today)]
compact_grn = len(td) > 6   # 場次多時 ✅組縮一行

lines = [f"📅 法說提醒 {today.month}/{today.day}（週{WD[today.weekday()]}）", ""]
if stale_warn:
    lines.insert(1, stale_warn)
if red:
    lines.append("🔴 財報未公布 — 當天出數字，重點盯")
    for e in red:
        lines.append(line1(e))
        lines.append(line2(e, False))
if grn:
    if red: lines.append("")
    lines.append("✅ 財報已出 — 法說為展望說明")
    for e in grn:
        if compact_grn:
            lines.append(line1(e) + f" ─ {brokers(e) or '—'}" + ("" if e.get("reacted") else "｜可能補報告"))
        else:
            lines.append(line1(e))
            lines.append(line2(e, True))
if td and not red:
    lines.insert(2, "（今日場次皆已出財報）")
if not td:
    lines.append("今天沒有場次")
if tm:
    lines.append("──")
    tags = []
    for e in tm[:8]:
        tags.append(("🔴" if not is_filed(e, tomorrow) else "✅") + f"{e['code']} {e['name']}")
    lines.append(f"明日 {tomorrow.month}/{tomorrow.day}：" + "、".join(tags) + ("…" if len(tm) > 8 else ""))
lines.append("")
lines.append("💡 🔴場次 82% 目標價調整當天發；✅但尚無反應的場次可等補報告")
text = "\n".join(lines)
print(text)
send(text)
