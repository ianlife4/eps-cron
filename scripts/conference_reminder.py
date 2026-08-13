# -*- coding: utf-8 -*-
"""法說/財報 中午提醒 — 讀 aurora research-map/events.json,推 TG (@Ianke_eps_daily_bot)
資料來源:外資研究地圖每日 19:30 更新的官方場次(MOPS)。今天有場次才推,沒有就靜默。
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

now = datetime.datetime.utcnow() + datetime.timedelta(hours=8)   # TW time
today = now.date()
tomorrow = today + datetime.timedelta(days=1)

req = urllib.request.Request(EVENTS_URL, headers={"User-Agent": "Mozilla/5.0"})
data = json.load(urllib.request.urlopen(req, timeout=30))
evs = data.get("officialEvents", [])

def day_events(d):
    out = [e for e in evs if e.get("date") == d.isoformat()]
    out.sort(key=lambda e: e.get("time") or "99")
    return out

def fmt(e, with_time=True):
    t = (e.get("time") or "").strip()
    loc = (e.get("loc") or "").strip()
    online = any(k in loc for k in ("線上", "Webex", "WEBEX", "webex", "Teams", "TEAMS", "電話", "視訊"))
    loc_tag = "線上" if online else (loc[:10] if loc else "")
    br = "·".join(SHORT.get(b, b) for b in (e.get("reactors") or [])[:5])
    parts = []
    if with_time and t: parts.append(t)
    parts.append(f"{e['code']} {e['name']}")
    if loc_tag: parts.append(f"({loc_tag})")
    if br: parts.append(f"─ {br}")
    return " ".join(parts)

td, tm = day_events(today), day_events(tomorrow)
if not td and not tm:
    print("no events today/tomorrow, skip")
    sys.exit(0)

lines = [f"📅 法說提醒 {today.month}/{today.day}（週{WD[today.weekday()]}）"]
if td:
    lines += [fmt(e) for e in td]
else:
    lines.append("今天沒有場次")
if tm:
    lines.append("──")
    lines.append(f"明日 {tomorrow.month}/{tomorrow.day}：" + "、".join(f"{e['code']} {e['name']}" for e in tm[:8])
                 + ("…" if len(tm) > 8 else ""))
lines.append("")
lines.append("💡 法說當天到隔天是外資報告高峰（歷史 82% 目標價調整當天發）")
text = "\n".join(lines)

payload = urllib.parse.urlencode({"chat_id": CHAT, "text": text}).encode()
r = urllib.request.urlopen(
    urllib.request.Request(f"https://api.telegram.org/bot{TOKEN}/sendMessage", data=payload), timeout=30)
resp = json.load(r)
print("sent:", resp.get("ok"), "msg_id:", resp.get("result", {}).get("message_id"))
