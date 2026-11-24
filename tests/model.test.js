// node tests/model.test.js — loads the .pragma-library JS files under node.
const fs = require("fs"), path = require("path"), assert = require("assert/strict")
function load(p) {
  const src = fs.readFileSync(path.join(__dirname, "..", p), "utf8").replace(/^\.pragma library/m, "")
  const module = { exports: {} }
  new Function("module", src)(module)
  return module.exports
}
const M = load("plugins/steve.meetings/Model.js")
const C = load("plugins/steve.teams-chat/Model.js")
const now = new Date(2026, 10, 24, 10, 0).getTime(), m = 60000
const ev = (s, a, b, x = {}) => ({ subject: s, start: now + a * m, end: now + b * m, response: "accepted", showAs: "busy", joinUrl: "https://teams.microsoft.com/l/x", ...x })
const st = (events, extra = {}) => ({ auth: "ok", updated: now, events, chats: [], presence: {}, ...extra })

let v = M.compute(null, now); assert.equal(v.status, "offline"); assert.equal(v.slides.length, 0)
v = M.compute(st([], { auth: "needed" }), now); assert.equal(v.status, "auth")
v = M.compute(st([]), now); assert.equal(v.status, "none"); assert.match(v.tooltip, /No meetings/)

// in a meeting with a back-to-back next one → two slides, next is urgent
v = M.compute(st([ev("Standup", -20, 10), ev("Client", 3, 30)]), now)
assert.equal(v.status, "current"); assert.equal(v.slides.length, 2)
assert.equal(v.slides[0].text, "Standup · 10m left"); assert.equal(v.slides[0].live, true)
assert.equal(v.slides[1].level, "urgent"); assert.match(v.slides[1].text, /^Next 10:03 Client · in 3m$/)
assert.equal(v.joinUrl, "https://teams.microsoft.com/l/x")

// next far away while busy → dim
v = M.compute(st([ev("Standup", -20, 10), ev("Later", 200, 230)]), now)
assert.equal(v.slides[1].level, "dim")

// prefers accepted over tentative when overlapping
v = M.compute(st([ev("Tent", -5, 30, { response: "tentativelyAccepted" }), ev("Real", -30, 30)]), now)
assert.match(v.slides[0].text, /^Real \(\+1\)/)

// upcoming only: thresholds
assert.equal(M.compute(st([ev("A", 4, 30)]), now).slides[0].level, "urgent")
assert.equal(M.compute(st([ev("A", 12, 30)]), now).slides[0].level, "accent")
assert.equal(M.compute(st([ev("A", 40, 60)]), now).slides[0].level, "normal")
v = M.compute(st([ev("A", 150, 160)]), now); assert.equal(v.slides[0].level, "dim"); assert.equal(v.slides[0].text, "12:30 A")
v = M.compute(st([ev("A", 24 * 60, 24 * 60 + 30)]), now); assert.equal(v.slides[0].text, "Wed 10:00 A")
assert.equal(v.joinUrl, "")  // too far out to offer join

// ad-hoc call from presence with no calendar event
v = M.compute(st([], { presence: { activity: "InACall" } }), now)
assert.equal(v.slides[0].text, "In a call"); assert.equal(v.status, "current")

// join URL allowlist rejects junk and shell metacharacters
assert.equal(M.safeJoin("https://evil.example/teams.microsoft.com"), "")
assert.equal(M.safeJoin("https://teams.microsoft.com/x'; rm -rf ~"), "")
assert.equal(M.safeJoin("javascript:alert(1)"), "")
assert.equal(M.safeJoin("https://us02web.zoom.us/j/123"), "https://us02web.zoom.us/j/123")
assert.equal(M.dur(1), "1m"); assert.equal(M.dur(80 * m), "1h 20m"); assert.equal(M.dur(120 * m), "2h")

// stale + error surfaced in tooltip
v = M.compute(st([], { updated: { calendar: now - 30 * m }, errors: { calendar: "boom" } }), now)
assert.match(v.tooltip, /calendar stale/); assert.match(v.tooltip, /boom/)
v = M.compute(st([], { updated: { calendar: now - 1 * m }, errors: { chats: "chat-only" } }), now)
assert.doesNotMatch(v.tooltip, /stale|chat-only/)

// current meeting stays visible (accent) however long is left; urgent near the end
assert.equal(M.compute(st([ev("Long", -10, 120)]), now).slides[0].level, "accent")
assert.equal(M.compute(st([ev("Ending", -50, 3)]), now).slides[0].level, "urgent")

// chat
const msg = (id, ago, x = {}) => ({ chatId: id, from: "Alex Example", text: "hi", time: now - ago * m, unread: true, ...x })
let c = C.compute({ chats: [msg("a", 1), msg("b", 30, { unread: false, topic: "Ops" })] }, now)
assert.equal(c.msg.chatId, "a"); assert.equal(c.hidden, false); assert.equal(c.dim, false); assert.equal(c.unreadCount, 1)
assert.match(c.tooltip, /\(Ops\)/)
assert.equal(C.compute({ chats: [msg("a", 1)] }, now, { dismissedKey: c.key }).hidden, true)
assert.equal(C.compute({ chats: [msg("a", 1, { unread: false })] }, now).dim, true)
assert.equal(C.compute({ chats: [msg("a", 20)] }, now).dim, true)
assert.equal(C.compute({ chats: [msg("a", 90)] }, now, { hideAfterMinutes: 60 }).hidden, true)
assert.equal(C.compute(null, now).hidden, true)
assert.equal(C.hue("Alex"), C.hue("Alex"))
c = C.compute({ chats: [msg("a", 1)], updated: { chats: now - 30 * m }, errors: { chats: "throttled" } }, now)
assert.match(c.tooltip, /chats stale/); assert.match(c.tooltip, /throttled/)
console.log("all model tests passed")
