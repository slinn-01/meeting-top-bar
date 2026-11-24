#!/usr/bin/env -S uv run --quiet --script
# /// script
# requires-python = ">=3.11"
# dependencies = ["msal>=1.28", "requests>=2.31"]
# ///
"""m365bar — poll Microsoft Graph for calendar, Teams chats, and presence,
and write a small JSON state file for the Omarchy bar widgets to read.

    m365bar.py login     device-code sign-in (once; refresh token is cached)
    m365bar.py run       poll forever (what the systemd user service runs)
    m365bar.py once      poll once and print the state
    m365bar.py logout    forget the cached tokens

SIGUSR1 to a running `run` forces an immediate refresh (bar middle-click).
"""
from __future__ import annotations

import html
import json
import os
import re
import signal
import stat
import sys
import threading
import time
import tomllib
from datetime import datetime, timedelta, timezone
from pathlib import Path

import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["User.Read", "Calendars.Read", "Chat.Read", "Presence.Read"]

CONFIG_PATH = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "m365bar" / "config.toml"
STATE_DIR = Path(os.environ.get("XDG_STATE_HOME", Path.home() / ".local" / "state")) / "m365bar"
TOKEN_CACHE = STATE_DIR / "token_cache.json"
RUNTIME_DIR = Path(os.environ.get("XDG_RUNTIME_DIR", f"/run/user/{os.getuid()}")) / "m365bar"
OUT_PATH = RUNTIME_DIR / "state.json"

DEFAULTS = {
    "client_id": "",
    "tenant": "organizations",
    "lookahead_hours": 18,
    "calendar_interval": 60,
    "chat_interval": 20,
    "presence_interval": 30,
    "include_free": False,       # include events marked Free (focus blocks etc.)
    "include_tentative": True,
    "chat_history": 8,           # messages kept for the tooltip
}

MAX_TEXT = 400
MAX_NAME = 60


class AuthNeeded(Exception):
    pass


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, file=sys.stderr, flush=True)


def load_config() -> dict:
    cfg = dict(DEFAULTS)
    if CONFIG_PATH.exists():
        with CONFIG_PATH.open("rb") as f:
            cfg.update(tomllib.load(f))
    if not cfg["client_id"]:
        sys.exit(f"m365bar: set client_id in {CONFIG_PATH} (see README)")
    return cfg


# ---------------------------------------------------------------- auth

class Auth:
    def __init__(self, cfg: dict):
        STATE_DIR.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.cache = msal.SerializableTokenCache()
        if TOKEN_CACHE.exists():
            self.cache.deserialize(TOKEN_CACHE.read_text())
        self.app = msal.PublicClientApplication(
            cfg["client_id"],
            authority=f"https://login.microsoftonline.com/{cfg['tenant']}",
            token_cache=self.cache,
        )

    def _save(self):
        if self.cache.has_state_changed:
            tmp = TOKEN_CACHE.with_suffix(".tmp")
            fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, stat.S_IRUSR | stat.S_IWUSR)
            with os.fdopen(fd, "w") as f:
                f.write(self.cache.serialize())
            os.replace(tmp, TOKEN_CACHE)

    def token(self) -> str:
        accounts = self.app.get_accounts()
        result = self.app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
        self._save()
        if not result or "access_token" not in result:
            raise AuthNeeded((result or {}).get("error_description", "no cached account"))
        return result["access_token"]

    def login(self):
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            sys.exit(f"device flow failed: {flow}")
        print(flow["message"], flush=True)
        result = self.app.acquire_token_by_device_flow(flow)
        self._save()
        if "access_token" not in result:
            sys.exit(f"login failed: {result.get('error_description')}")
        print(f"signed in as {result.get('id_token_claims', {}).get('preferred_username', '?')}")

    def logout(self):
        TOKEN_CACHE.unlink(missing_ok=True)


# ---------------------------------------------------------------- graph

class Graph:
    def __init__(self, auth: Auth):
        self.auth = auth
        self.s = requests.Session()

    def get(self, path: str, params=None, headers=None) -> dict:
        h = {"Authorization": f"Bearer {self.auth.token()}"}
        h.update(headers or {})
        r = self.s.get(GRAPH + path, params=params, headers=h, timeout=20)
        if r.status_code == 401:
            raise AuthNeeded(r.text[:200])
        if r.status_code == 429:
            retry = int(r.headers.get("Retry-After", "30"))
            raise RuntimeError(f"throttled, retry after {retry}s")
        r.raise_for_status()
        return r.json()


def ms(dt_str: str | None) -> int | None:
    """Graph dateTime (UTC, no offset, 7-digit fraction) -> epoch ms."""
    if not dt_str:
        return None
    s = dt_str.rstrip("Z")
    if "." in s:
        head, frac = s.split(".", 1)
        s = f"{head}.{frac[:6]}"
    dt = datetime.fromisoformat(s)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return int(dt.timestamp() * 1000)


def clip(s, n=MAX_TEXT) -> str:
    s = str(s or "")
    return s if len(s) <= n else s[: n - 1] + "…"


JOIN_HOSTS = re.compile(
    r"^https://([a-z0-9-]+\.)*(teams\.microsoft\.com|teams\.live\.com|zoom\.us|meet\.google\.com|webex\.com|gotomeeting\.com)(/|$)",
    re.I,
)


def safe_url(u) -> str:
    u = str(u or "")
    return u if JOIN_HOSTS.match(u) and len(u) < 2048 and not re.search(r"[\s'\"`\\]", u) else ""


def fetch_calendar(g: Graph, cfg: dict) -> list[dict]:
    now = datetime.now(timezone.utc)
    params = {
        "startDateTime": (now - timedelta(hours=12)).isoformat(),
        "endDateTime": (now + timedelta(hours=cfg["lookahead_hours"])).isoformat(),
        "$select": "subject,start,end,isAllDay,isCancelled,showAs,responseStatus,"
                   "isOnlineMeeting,onlineMeeting,location,organizer,webLink,isOrganizer",
        "$orderby": "start/dateTime",
        "$top": "100",
    }
    data = g.get("/me/calendarView", params=params, headers={"Prefer": 'outlook.timezone="UTC"'})
    out = []
    for e in data.get("value", []):
        if e.get("isCancelled") or e.get("isAllDay"):
            continue
        response = (e.get("responseStatus") or {}).get("response", "none")
        show_as = e.get("showAs", "busy")
        if response == "declined":
            continue
        if show_as == "free" and not cfg["include_free"]:
            continue
        if response == "tentativelyAccepted" and not cfg["include_tentative"]:
            continue
        join = safe_url((e.get("onlineMeeting") or {}).get("joinUrl")) or safe_url((e.get("location") or {}).get("locationUri"))
        out.append({
            "subject": clip(e.get("subject") or "(no title)", 200),
            "start": ms((e.get("start") or {}).get("dateTime")),
            "end": ms((e.get("end") or {}).get("dateTime")),
            "showAs": show_as,
            "response": response,
            "online": bool(join),
            "joinUrl": join,
            "location": clip((e.get("location") or {}).get("displayName"), 120),
            "organizer": clip(((e.get("organizer") or {}).get("emailAddress") or {}).get("name"), MAX_NAME),
        })
    return out


_TAG_ATTACH = re.compile(r"<attachment[^>]*>.*?</attachment>|<attachment[^>]*/?>", re.I | re.S)
_TAG_EMOJI = re.compile(r"<emoji[^>]*\balt=\"([^\"]*)\"[^>]*>(?:.*?</emoji>)?", re.I | re.S)
_TAG_IMG = re.compile(r"<img[^>]*>", re.I)
_TAG_BREAK = re.compile(r"<(br|/p|/div|/li)[^>]*>", re.I)
_TAG_ANY = re.compile(r"<[^>]+>")


def plain(body: dict | None) -> str:
    body = body or {}
    text = body.get("content") or ""
    if (body.get("contentType") or "").lower() == "html":
        text = _TAG_ATTACH.sub(" [attachment] ", text)
        text = _TAG_EMOJI.sub(lambda m: m.group(1), text)
        text = _TAG_IMG.sub(" [image] ", text)
        text = _TAG_BREAK.sub(" ", text)
        text = _TAG_ANY.sub("", text)
        text = html.unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    return clip(text)


def fetch_chats(g: Graph, cfg: dict, my_id: str) -> list[dict]:
    params = {
        "$expand": "lastMessagePreview",
        "$orderby": "lastMessagePreview/createdDateTime desc",
        "$top": "25",
    }
    data = g.get("/me/chats", params=params)
    out = []
    for c in data.get("value", []):
        m = c.get("lastMessagePreview") or {}
        if not m or m.get("isDeleted") or m.get("messageType") != "message":
            continue
        user = ((m.get("from") or {}).get("user") or {})
        if not user or user.get("id") == my_id:
            continue
        text = plain(m.get("body"))
        if not text:
            continue
        created = ms(m.get("createdDateTime"))
        read_at = ms(((c.get("viewpoint") or {}).get("lastMessageReadDateTime")))
        web = str(c.get("webUrl") or "")
        out.append({
            "chatId": clip(c.get("id"), 200),
            "chatType": c.get("chatType", ""),
            "topic": clip(c.get("topic"), MAX_NAME),
            "from": clip(user.get("displayName") or "Someone", MAX_NAME),
            "text": text,
            "time": created,
            "unread": bool(created and (read_at is None or read_at < created)),
            "webUrl": web if web.startswith("https://teams.microsoft.com/") and not re.search(r"[\s'\"`\\]", web) else "",
        })
        if len(out) >= cfg["chat_history"]:
            break
    out.sort(key=lambda x: x["time"] or 0, reverse=True)
    return out


def fetch_presence(g: Graph) -> dict:
    p = g.get("/me/presence")
    return {"availability": p.get("availability", ""), "activity": p.get("activity", "")}


# ---------------------------------------------------------------- state

def write_state(state: dict):
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True, mode=0o700)
    tmp = OUT_PATH.with_suffix(".tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, stat.S_IRUSR | stat.S_IWUSR)
    with os.fdopen(fd, "w") as f:
        json.dump(state, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, OUT_PATH)


def run(cfg: dict, once=False):
    auth = Auth(cfg)
    g = Graph(auth)
    wake = threading.Event()
    signal.signal(signal.SIGUSR1, lambda *_: wake.set())

    state = {"version": 1, "auth": "ok", "error": "", "updated": 0,
             "events": [], "chats": [], "presence": {}}
    last = {"calendar": 0.0, "chat": 0.0, "presence": 0.0}
    my_id = ""
    backoff = 0

    while True:
        forced = wake.is_set()
        wake.clear()
        now = time.monotonic()
        try:
            if not my_id:
                my_id = g.get("/me", params={"$select": "id"})["id"]
            if forced or now - last["calendar"] >= cfg["calendar_interval"]:
                state["events"] = fetch_calendar(g, cfg)
                last["calendar"] = now
            if forced or now - last["chat"] >= cfg["chat_interval"]:
                state["chats"] = fetch_chats(g, cfg, my_id)
                last["chat"] = now
            if forced or now - last["presence"] >= cfg["presence_interval"]:
                try:
                    state["presence"] = fetch_presence(g)
                except requests.HTTPError as e:  # presence is optional
                    log("presence:", e)
                last["presence"] = now
            state.update(auth="ok", error="", updated=int(time.time() * 1000))
            backoff = 0
        except AuthNeeded as e:
            log("auth needed:", e)
            state.update(auth="needed", error="Sign-in needed: run m365bar login")
            my_id = ""
            # `m365bar login` runs in another process; pick up its cache.
            auth = Auth(cfg)
            g.auth = auth
            backoff = 15
        except Exception as e:  # network blips: keep last good data, flag it
            log("error:", e)
            state["error"] = clip(str(e), 200)
            backoff = min(300, max(15, backoff * 2))
        write_state(state)
        if once:
            print(json.dumps(state, indent=2, ensure_ascii=False))
            return
        wake.wait(timeout=backoff or 5)


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "run"
    cfg = load_config()
    if cmd == "login":
        Auth(cfg).login()
    elif cmd == "logout":
        Auth(cfg).logout()
    elif cmd == "once":
        run(cfg, once=True)
    elif cmd == "run":
        run(cfg)
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
