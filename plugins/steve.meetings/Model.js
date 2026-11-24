.pragma library

// Pure view logic for the meetings bar widget. No QML types in here so it can
// be exercised with node (see tests/model.test.js).

var CALL_ACTIVITIES = ["InACall", "InAConferenceCall", "InAMeeting", "Presenting"]
var JOIN_RE = /^https:\/\/([a-z0-9-]+\.)*(teams\.microsoft\.com|teams\.live\.com|zoom\.us|meet\.google\.com|webex\.com|gotomeeting\.com)(\/[^\s'"`\\]*)?$/i
var MIN = 60000

function pad(n) { return n < 10 ? "0" + n : "" + n }

function clock(ms) {
  var d = new Date(ms)
  return pad(d.getHours()) + ":" + pad(d.getMinutes())
}

function sameDay(a, b) {
  var x = new Date(a), y = new Date(b)
  return x.getFullYear() === y.getFullYear() && x.getMonth() === y.getMonth() && x.getDate() === y.getDate()
}

var DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]

function when(ms, now) {
  return sameDay(ms, now) ? clock(ms) : DAYS[new Date(ms).getDay()] + " " + clock(ms)
}

// 3m, 45m, 1h, 1h 20m — always rounded up so "in 0m" never shows.
function dur(ms) {
  var m = Math.max(1, Math.ceil(ms / MIN))
  if (m < 60) return m + "m"
  var h = Math.floor(m / 60), r = m % 60
  return r ? h + "h " + r + "m" : h + "h"
}

function safeJoin(url) {
  return typeof url === "string" && url.length < 2048 && JOIN_RE.test(url) ? url : ""
}

function rank(e) {
  // Prefer meetings I accepted / organize over tentative ones.
  if (e.response === "organizer" || e.response === "accepted") return 2
  if (e.response === "tentativelyAccepted") return 0
  return 1
}

function validEvents(state) {
  var list = state && Array.isArray(state.events) ? state.events : []
  return list.filter(function(e) {
    return e && typeof e.start === "number" && typeof e.end === "number" && e.end > e.start
  })
}

function current(events, now) {
  return events.filter(function(e) { return e.start <= now && now < e.end })
    .sort(function(a, b) { return rank(b) - rank(a) || b.start - a.start })
}

function upcoming(events, now) {
  return events.filter(function(e) { return e.start > now })
    .sort(function(a, b) { return a.start - b.start })
}

// Returns { status, slides: [{kind, subject, text, level, live}], tooltip, joinUrl }
// level: "urgent" | "accent" | "normal" | "dim"
function compute(state, now, opts) {
  opts = opts || {}
  var soon = (opts.soonMinutes || 15) * MIN
  var urgent = (opts.urgentMinutes || 5) * MIN
  var joinLead = (opts.joinLeadMinutes || 10) * MIN
  var staleAfter = (opts.staleMinutes || 10) * MIN

  if (!state) {
    return { status: "offline", slides: [], tooltip: "m365bar daemon not running", joinUrl: "" }
  }
  if (state.auth === "needed") {
    return {
      status: "auth",
      slides: [{ kind: "auth", subject: "", text: "M365 sign-in needed", level: "urgent", live: false }],
      tooltip: "Microsoft 365 sign-in needed — click to sign in",
      joinUrl: ""
    }
  }

  var events = validEvents(state)
  var cur = current(events, now)
  var next = upcoming(events, now)
  var presence = state.presence || {}
  var inCall = CALL_ACTIVITIES.indexOf(presence.activity) >= 0
  var slides = []
  var joinUrl = ""

  if (cur.length) {
    var c = cur[0]
    var more = cur.length > 1 ? " (+" + (cur.length - 1) + ")" : ""
    var left = c.end - now
    slides.push({
      kind: "current", subject: c.subject,
      text: c.subject + more + " · " + dur(left) + " left",
      level: left <= urgent ? "urgent" : "accent",
      live: true
    })
    joinUrl = safeJoin(c.joinUrl)
  } else if (inCall) {
    slides.push({ kind: "call", subject: "", text: presence.activity === "Presenting" ? "Presenting" : "In a call", level: "accent", live: true })
  }

  if (next.length) {
    var n = next[0]
    var until = n.start - now
    var level = until <= urgent ? "urgent" : until <= soon ? "accent" : until <= 60 * MIN ? "normal" : "dim"
    var text = until <= 60 * MIN
      ? clock(n.start) + " " + n.subject + " · in " + dur(until)
      : when(n.start, now) + " " + n.subject
    if (slides.length) {
      // While busy, the "next" slide is a heads-up; back-to-back gets loud.
      text = "Next " + text
      if (until > soon) level = "dim"
    }
    slides.push({ kind: "next", subject: n.subject, text: text, level: level, live: false })
    if (!joinUrl && until <= joinLead) joinUrl = safeJoin(n.joinUrl)
  }

  var lines = []
  if (cur.length) {
    lines.push("Now")
    cur.forEach(function(e) { lines.push("  " + clock(e.start) + "–" + clock(e.end) + "  " + e.subject + tag(e)) })
  } else if (inCall) {
    lines.push("Now: " + presence.activity)
  }
  if (next.length) {
    if (lines.length) lines.push("")
    lines.push("Upcoming")
    next.slice(0, 8).forEach(function(e) {
      lines.push("  " + when(e.start, now) + "–" + clock(e.end) + "  " + e.subject + tag(e))
    })
  }
  if (!lines.length) lines.push("No meetings in the next " + (opts.lookaheadHours || 18) + "h")
  if (state.updated && now - state.updated > staleAfter) lines.push("", "⚠ data stale since " + clock(state.updated))
  if (state.error) lines.push("", "⚠ " + state.error)
  lines.push("", "Click: calendar · Right: join" + (joinUrl ? "" : " (none)") + " · Middle: refresh")

  return {
    status: slides.length ? (cur.length || inCall ? "current" : "upcoming") : "none",
    slides: slides,
    tooltip: lines.join("\n"),
    joinUrl: joinUrl
  }
}

function tag(e) {
  var t = []
  if (e.response === "tentativelyAccepted" || e.showAs === "tentative") t.push("tentative")
  if (e.response === "notResponded") t.push("not responded")
  if (e.online) t.push("online")
  else if (e.location) t.push(e.location)
  return t.length ? "  [" + t.join(", ") + "]" : ""
}

if (typeof module !== "undefined") module.exports = { compute: compute, dur: dur, safeJoin: safeJoin }
