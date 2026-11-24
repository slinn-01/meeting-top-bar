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
    m365bar.py logout    stop the service and forget the cached tokens

SIGUSR1 to a running `run` forces an immediate refresh (bar middle-click).
"""
from __future__ import annotations

import html
import json
import os
import random
import re
import signal
import stat
import subprocess
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
BOUNDS = {
    "lookahead_hours": (1, 168),
    "calendar_interval": (15, 3600),
    "chat_interval": (5, 3600),
    "presence_interval": (10, 3600),
    "chat_history": (1, 50),
}

MAX_TEXT = 400
MAX_NAME = 60
MAX_PAGES = 3
HEARTBEAT = 120  # rewrite unchanged state this often so widgets can judge freshness


class AuthNeeded(Exception):
    pass


class Throttled(Exception):
    def __init__(self, retry_after: int):
        super().__init__(f"throttled, retry after {retry_after}s")
        self.retry_after = retry_after


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, file=sys.stderr, flush=True)


def load_config() -> dict:
    cfg = dict(DEFAULTS)
    if CONFIG_PATH.exists():
        with CONFIG_PATH.open("rb") as f:
            cfg.update(tomllib.load(f))
    problems = []
    if not isinstance(cfg["client_id"], str) or not cfg["client_id"]:
        problems.append("client_id must be set (see README)")
    for key, (lo, hi) in BOUNDS.items():
        v = cfg[key]
        if isinstance(v, bool) or not isinstance(v, int) or not lo <= v <= hi:
            problems.append(f"{key} must be an integer in {lo}..{hi}")
    for key in ("include_free", "include_tentative"):
        if not isinstance(cfg[key], bool):
            problems.append(f"{key} must be true or false")
    if problems:
        sys.exit(f"m365bar: bad {CONFIG_PATH}:\n  " + "\n  ".join(problems))
    return cfg


def runtime_state_path() -> Path:
    runtime = os.environ.get("XDG_RUNTIME_DIR")
    if not runtime:
        sys.exit("m365bar: XDG_RUNTIME_DIR is not set")
    return Path(runtime) / "m365bar" / "state.json"


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
            write_private(TOKEN_CACHE, self.cache.serialize())

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


def logout():
    # The service holds tokens in memory, so it must be down before the cache goes.
    subprocess.run(["systemctl", "--user", "stop", "m365bar.service"], check=False)
    state = subprocess.run(["systemctl", "--user", "is-active", "m365bar.service"],
                           capture_output=True, text=True).stdout.strip()
    if state not in ("inactive", "failed", "unknown"):
        sys.exit(f"m365bar: service is still {state or 'in an unknown state'}; tokens left in place")
    TOKEN_CACHE.unlink(missing_ok=True)
    print("m365bar service stopped and tokens removed; `m365bar login` then start the service again")


# ---------------------------------------------------------------- graph

class Graph:
    def __init__(self, auth: Auth):
        self.auth = auth
        self.s = requests.Session()

    def get(self, url: str, params=None, headers=None) -> dict:
        if not url.startswith(GRAPH + "/"):
            url = GRAPH + url
        h = {"Authorization": f"Bearer {self.auth.token()}"}
        h.update(headers or {})
        r = self.s.get(url, params=params, headers=h, timeout=20)
        if r.status_code == 401:
            raise AuthNeeded(r.text[:200])
        if r.status_code == 429:
            raise Throttled(int(r.headers.get("Retry-After", "60")))
        r.raise_for_status()
        return r.json()

    def items(self, path: str, params=None, headers=None):
        """Yield collection items, following @odata.nextLink up to MAX_PAGES."""
        data = self.get(path, params=params, headers=headers)
        for page in range(MAX_PAGES):
            yield from data.get("value", [])
            nxt = data.get("@odata.nextLink")
            if page + 1 == MAX_PAGES or not isinstance(nxt, str) or not nxt.startswith(GRAPH + "/"):
                return
            data = self.get(nxt, headers=headers)


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


UNSAFE_URL_CHARS = re.compile(r"[\s'\"`\\<>]")
JOIN_HOSTS = re.compile(
    r"^https://([a-z0-9-]+\.)*(teams\.microsoft\.com|teams\.live\.com|zoom\.us|meet\.google\.com|webex\.com|gotomeeting\.com)(/|$)",
    re.I,
)


def safe_url(u, host_re=JOIN_HOSTS) -> str:
    u = str(u or "")
    return u if len(u) < 2048 and host_re.match(u) and not UNSAFE_URL_CHARS.search(u) else ""


def fetch_calendar(g: Graph, cfg: dict, my_id: str) -> list[dict]:
    now = datetime.now(timezone.utc)
    params = {
        "startDateTime": (now - timedelta(hours=12)).isoformat(),
        "endDateTime": (now + timedelta(hours=cfg["lookahead_hours"])).isoformat(),
        "$select": "subject,start,end,isAllDay,isCancelled,showAs,responseStatus,onlineMeeting,location",
        "$orderby": "start/dateTime",
        "$top": "100",
    }
    out = []
    for e in g.items("/me/calendarView", params=params, headers={"Prefer": 'outlook.timezone="UTC"'}):
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
        location = e.get("location") or {}
        join = safe_url((e.get("onlineMeeting") or {}).get("joinUrl")) or safe_url(location.get("locationUri"))
        out.append({
            "subject": clip(e.get("subject") or "(no title)", 200),
            "start": ms((e.get("start") or {}).get("dateTime")),
            "end": ms((e.get("end") or {}).get("dateTime")),
            "showAs": show_as,
            "response": response,
            "online": bool(join),
            "joinUrl": join,
            "location": clip(location.get("displayName"), 120),
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


TEAMS_CHAT = re.compile(r"^https://teams\.microsoft\.com/")


def fetch_chats(g: Graph, cfg: dict, my_id: str) -> list[dict]:
    params = {
        "$expand": "lastMessagePreview",
        "$orderby": "lastMessagePreview/createdDateTime desc",
        "$top": "25",
    }
    out = []
    for c in g.items("/me/chats", params=params):
        m = c.get("lastMessagePreview") or {}
        if not m or m.get("isDeleted") or m.get("messageType") != "message":
            continue
        user = (m.get("from") or {}).get("user") or {}
        if not user or user.get("id") == my_id:
            continue
        text = plain(m.get("body"))
        if not text:
            continue
        created = ms(m.get("createdDateTime"))
        read_at = ms((c.get("viewpoint") or {}).get("lastMessageReadDateTime"))
        out.append({
            "chatId": clip(c.get("id"), 200),
            "topic": clip(c.get("topic"), MAX_NAME),
            "from": clip(user.get("displayName") or "Someone", MAX_NAME),
            "text": text,
            "time": created,
            "unread": bool(created and (read_at is None or read_at < created)),
            "webUrl": safe_url(c.get("webUrl"), TEAMS_CHAT),
        })
        if len(out) >= cfg["chat_history"]:
            break
    out.sort(key=lambda x: x["time"] or 0, reverse=True)
    return out


def fetch_presence(g: Graph, cfg: dict, my_id: str) -> dict:
    return {"activity": g.get("/me/presence").get("activity", "")}


# ---------------------------------------------------------------- state

def write_private(path: Path, text: str):
    tmp = path.with_suffix(".tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, stat.S_IRUSR | stat.S_IWUSR)
    with os.fdopen(fd, "w") as f:
        f.write(text)
    os.replace(tmp, path)


# name -> (fetcher, state key, interval config key)
FEEDS = {
    "calendar": (fetch_calendar, "events", "calendar_interval"),
    "chats": (fetch_chats, "chats", "chat_interval"),
    "presence": (fetch_presence, "presence", "presence_interval"),
}


def run(cfg: dict, once=False):
    out_path = runtime_state_path()
    out_path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    auth = Auth(cfg)
    g = Graph(auth)
    wake = threading.Event()
    signal.signal(signal.SIGUSR1, lambda *_: wake.set())

    data = {"events": [], "chats": [], "presence": {}}
    # "next" is the normal schedule (a forced refresh may skip it); "hold" is a
    # server Retry-After deadline that nothing skips.
    feeds = {name: {"next": 0.0, "hold": 0.0, "updated": 0, "error": "", "fails": 0} for name in FEEDS}
    auth_state = "ok"
    my_id = ""
    last_content, last_write = None, 0.0

    while True:
        forced = wake.is_set()
        wake.clear()
        try:
            if not my_id:
                my_id = g.get("/me", params={"$select": "id"})["id"]
            # Each feed fails, backs off, and honors Retry-After on its own so
            # one broken endpoint never starves the others.
            for name, (fetch, key, interval_key) in FEEDS.items():
                f = feeds[name]
                now = time.monotonic()
                if now < f["hold"] or (not forced and now < f["next"]):
                    continue
                try:
                    data[key] = fetch(g, cfg, my_id)
                    f.update(updated=int(time.time() * 1000), error="", fails=0, next=now + cfg[interval_key])
                except Throttled as e:
                    log(f"{name}:", e)
                    f["hold"] = f["next"] = now + e.retry_after + random.uniform(1, 5)
                    f["error"] = str(e)
                except AuthNeeded:
                    raise
                except Exception as e:  # network blips: keep last good data, flag it
                    log(f"{name}:", e)
                    f["fails"] += 1
                    f.update(error=clip(str(e), 200), next=now + min(300, 15 * 2 ** (f["fails"] - 1)))
            auth_state = "ok"
        except AuthNeeded as e:
            log("auth needed:", e)
            auth_state, my_id = "needed", ""
            # `m365bar login` runs in another process; pick up its cache next time.
            auth = Auth(cfg)
            g.auth = auth
            for f in feeds.values():
                f["next"] = time.monotonic() + 15
        except Exception as e:  # the /me lookup itself failed
            log("me:", e)
            for f in feeds.values():
                f.update(error=clip(str(e), 200), next=time.monotonic() + 30)

        state = {
            "version": 2,
            "auth": auth_state,
            **data,
            "errors": {n: f["error"] for n, f in feeds.items()},
        }
        content = json.dumps(state, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
        if content != last_content or time.monotonic() - last_write >= HEARTBEAT or once:
            state["updated"] = {n: f["updated"] for n, f in feeds.items()}
            write_private(out_path, json.dumps(state, ensure_ascii=False, separators=(",", ":")))
            last_content, last_write = content, time.monotonic()
        if once:
            print(json.dumps(state, indent=2, ensure_ascii=False))
            return
        next_due = min(f["next"] for f in feeds.values())
        wake.wait(timeout=min(HEARTBEAT, max(1.0, next_due - time.monotonic())))


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "run"
    if cmd not in ("login", "logout", "once", "run"):
        sys.exit(__doc__)
    if cmd == "logout":
        return logout()
    cfg = load_config()
    if cmd == "login":
        Auth(cfg).login()
    else:
        run(cfg, once=(cmd == "once"))


if __name__ == "__main__":
    main()
