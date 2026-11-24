#!/usr/bin/env python3
"""Write a fake state file so the widgets can be checked without Graph.
   usage: demo-state.py [incall|soon|later|auth|empty]"""
import json, os, sys, time
mode = sys.argv[1] if len(sys.argv) > 1 else "incall"
now = int(time.time() * 1000); M = 60_000
ev = lambda s, a, b, **k: {"subject": s, "start": now + a * M, "end": now + b * M, "showAs": "busy",
                           "response": "accepted", "online": True,
                           "joinUrl": "https://teams.microsoft.com/l/meetup-join/demo", "location": "", "organizer": "Demo", **k}
events = {
    "incall": [ev("Weekly standup", -20, 10), ev("Client sync <b>ACME</b>", 12, 42), ev("1:1 with Jess", 120, 150)],
    "soon":   [ev("Client sync <b>ACME</b>", 4, 34), ev("Design review", 90, 120, response="tentativelyAccepted")],
    "later":  [ev("Design review", 95, 125)],
    "auth": [], "empty": [],
}[mode]
chats = [
    {"chatId": "c1", "chatType": "oneOnOne", "topic": "", "from": "Alex Example",
     "text": "Did the deploy go out? <img src=x>", "time": now - 30_000, "unread": True,
     "webUrl": "https://teams.microsoft.com/l/chat/demo/0"},
    {"chatId": "c2", "chatType": "group", "topic": "Ops", "from": "Jess Q", "text": "deploy done", "time": now - 40 * M,
     "unread": False, "webUrl": ""},
] if mode != "empty" else []
state = {"version": 1, "auth": "needed" if mode == "auth" else "ok", "error": "", "updated": now,
         "events": events, "chats": chats,
         "presence": {"availability": "Busy", "activity": "InACall" if mode == "incall" else "Available"}}
d = os.path.join(os.environ.get("XDG_RUNTIME_DIR", "/tmp"), "m365bar"); os.makedirs(d, exist_ok=True)
tmp = os.path.join(d, "state.tmp"); json.dump(state, open(tmp, "w")); os.replace(tmp, os.path.join(d, "state.json"))
print("wrote", mode)
