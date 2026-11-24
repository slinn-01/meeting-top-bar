.pragma library

// Pure view logic for the Teams chat ticker (node-testable).

var MIN = 60000

function ago(ms, now) {
  var d = Math.max(0, now - ms)
  if (d < MIN) return "now"
  if (d < 60 * MIN) return Math.floor(d / MIN) + "m"
  if (d < 24 * 60 * MIN) return Math.floor(d / (60 * MIN)) + "h"
  return Math.floor(d / (24 * 60 * MIN)) + "d"
}

// Stable per-person hue, like Twitch name colors.
function hue(name) {
  var h = 0
  var s = String(name || "")
  for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0
  return (h % 360) / 360
}

function keyOf(m) { return m ? m.chatId + "|" + m.time : "" }

function chats(state) {
  var list = state && Array.isArray(state.chats) ? state.chats : []
  return list.filter(function(m) { return m && typeof m.text === "string" && typeof m.time === "number" })
}

// opts: { hideAfterMinutes (0 = never), dimAfterMinutes, dismissedKey }
// Returns { msg, key, dim, hidden, unreadCount, tooltip }
function compute(state, now, opts) {
  opts = opts || {}
  var list = chats(state)
  var msg = list.length ? list[0] : null
  var key = keyOf(msg)
  var age = msg ? now - msg.time : Infinity
  var hideAfter = (opts.hideAfterMinutes || 0) * MIN
  var dimAfter = (opts.dimAfterMinutes || 10) * MIN
  var unread = list.filter(function(m) { return m.unread }).length

  var hidden = !msg || key === opts.dismissedKey || (hideAfter > 0 && age > hideAfter)
  var lines = list.map(function(m) {
    var where = m.topic ? " (" + m.topic + ")" : ""
    return (m.unread ? "● " : "  ") + ago(m.time, now) + "  " + m.from + where + ": " + m.text
  })
  if (!lines.length) lines.push("No recent Teams messages")
  if (state && state.auth === "needed") lines.unshift("Microsoft 365 sign-in needed", "")
  lines.push("", "Click: Teams · Right: open chat · Middle: dismiss")

  return {
    msg: msg,
    key: key,
    dim: !msg || !msg.unread || age > dimAfter,
    hidden: hidden,
    unreadCount: unread,
    tooltip: lines.join("\n")
  }
}

if (typeof module !== "undefined") module.exports = { compute: compute, hue: hue, ago: ago }
