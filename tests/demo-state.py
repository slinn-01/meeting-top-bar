#!/usr/bin/env python3
"""Write a fake state file so the widgets can be checked without Graph.
   usage: demo-state.py [incall|soon|later|auth|empty]
   Stop m365bar.service first, or it will overwrite this within a few minutes."""
import json, os, sys, time
mode = sys.argv[1] if len(sys.argv) > 1 else "incall"
now = int(time.time() * 1000); M = 60_000
ev = lambda s, a, b, **k: {"subject": s, "start": now + a * M, "end": now + b * M, "showAs": "busy",
                           "response": "accepted", "online": True,
                           "joinUrl": "https://teams.microsoft.com/l/meetup-join/demo", "location": "", **k}
events = {
    "incall": [ev("Weekly standup", -20, 10), ev("Client sync <b>ACME</b>", 12, 42), ev("1:1 with Jess", 120, 150)],
    "soon":   [ev("Client sync <b>ACME</b>", 4, 34), ev("Design review", 90, 120, response="tentativelyAccepted")],
    "later":  [ev("Design review", 95, 125)],
    "auth": [], "empty": [],
}[mode]
chats = [
    {"chatId": "c1", "topic": "", "from": "Alex Example",
     "text": "Did the deploy go out? <img src=x>", "time": now - 30_000, "unread": True,
     "webUrl": "https://teams.microsoft.com/l/chat/demo/0"},
    {"chatId": "c2", "topic": "Ops", "from": "Jess Q", "text": "deploy done", "time": now - 40 * M,
     "unread": False, "webUrl": ""},
] if mode != "empty" else []
state = {"version": 2, "auth": "needed" if mode == "auth" else "ok",
         "events": events, "chats": chats,
         "presence": {"activity": "InACall" if mode == "incall" else "Available"},
         "errors": {"calendar": "", "chats": "", "presence": ""},
         "updated": {"calendar": now, "chats": now, "presence": now}}
runtime = os.environ.get("XDG_RUNTIME_DIR") or sys.exit("XDG_RUNTIME_DIR is not set")
d = os.path.join(runtime, "m365bar"); os.makedirs(d, mode=0o700, exist_ok=True)
tmp = os.path.join(d, "state.tmp"); json.dump(state, open(tmp, "w")); os.replace(tmp, os.path.join(d, "state.json"))
print("wrote", mode)
