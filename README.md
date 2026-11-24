# meeting-top-bar

Two Omarchy (Quattro / omarchy-shell) bar widgets fed by one Microsoft Graph poller:

- **`steve.meetings`** (centerBegin, beside the clock): current meeting with a pulsing dot and
  time left, crossfading every 6s to the next one. Countdown escalates: dim (>1h) → normal
  (<1h) → accent (<15m) → bold urgent (<5m). Teams presence "In a call" also counts as in-call
  for ad-hoc calls with no calendar event.
  Left: Outlook calendar web app · Right: join current/imminent online meeting · Middle: refresh.
- **`steve.teams-chat`** (right, before tray): latest Teams chat message from someone else, as a
  Twitch-chat-style pill with a per-person name color. New messages slide in; read or >10min
  old dims. Left: focus Teams · Right: open that chat · Middle: dismiss until next message.

```
Graph ──(m365bar daemon, systemd --user)──► $XDG_RUNTIME_DIR/m365bar/state.json ──► QML widgets
        calendarView 60s · chats 20s · presence 30s           (tmpfs, 0600)        (FileView)
```

## Why Graph instead of ICS

ICS covers the calendar only. Teams messages and "am I actually in a call" (presence) need
Graph anyway, so one delegated-auth poller does all three.

## One-time setup (Entra admin center)

1. **App registrations → New registration**: name `m365bar`, *single tenant*, no redirect URI.
2. **Authentication → Advanced settings → Allow public client flows: Yes** (device-code login).
3. **API permissions → Add → Microsoft Graph → Delegated**: `User.Read`, `Calendars.Read`,
   `Chat.Read`, `Presence.Read` → **Grant admin consent**.
4. Put the **Application (client) ID** and **Directory (tenant) ID** in
   `~/.config/m365bar/config.toml` (`client_id`, `tenant`).

Then:

```bash
m365bar login                                  # device code; token cached 0600 in ~/.local/state/m365bar
m365bar once | jq '.events[0], .chats[0]'      # sanity check
systemctl --user enable --now m365bar.service
```

Clicking the bar while it says "M365 sign-in needed" also runs `m365bar login`.

## Files

| Path | What |
|---|---|
| `daemon/m365bar.py` | uv inline-script poller (msal + requests) |
| `daemon/m365bar.service` | systemd user unit (`~/.local/bin/m365bar run`) |
| `plugins/steve.*/` | bar widgets; `Model.js` holds all view logic |
| `install.sh` | symlinks the daemon, copies plugins into `~/.config/omarchy/plugins` |
| `tests/model.test.js` | `node tests/model.test.js` |
| `tests/demo-state.py` | `incall \| soon \| later \| auth \| empty` fake state for UI checks |

Re-run `./install.sh` after editing the plugins (the shell hot-reloads them).

## Widget settings (`shell.json`, via `omarchy bar set <id> <key> <value>`)

- meetings: `soonMinutes` 15, `urgentMinutes` 5, `joinLeadMinutes` 10, `rotateSeconds` 6,
  `maxWidth` 380, `showWhenEmpty` true, `calendarCommand`
- teams-chat: `maxWidth` 600, `dimAfterMinutes` 10, `hideAfterMinutes` 0 (never),
  `pillOpacity` 0.85, `teamsCommand`

## Safety notes

Meeting subjects and chat text come from other people, including external senders. All text sinks
are `Text.PlainText` and the shared tooltip is plain text. Join and chat URLs are allowlisted
(Teams/Zoom/Meet/Webex/GoTo, https only, no shell metacharacters) in both the daemon and the QML
before being `Util.shellQuote`d into a launch command.

Before installing, back up your bar config: `cp ~/.config/omarchy/shell.json{,.bak}`.
