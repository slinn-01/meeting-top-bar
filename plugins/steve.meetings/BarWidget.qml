import QtQuick
import Quickshell
import Quickshell.Io
import qs.Commons
import qs.Commons as Commons
import qs.Ui
import "Model.js" as Model

// Current / next meeting from Microsoft 365, fed by the m365bar daemon
// ($XDG_RUNTIME_DIR/m365bar/state.json).
//
// In a meeting it alternates between the live meeting and the next one with a
// crossfade. Left click opens the Outlook calendar web app, right click joins
// the current (or imminent) online meeting, middle click forces a refresh.
BarWidget {
  id: root
  moduleName: "steve.meetings"

  // No /tmp fallback: a shared directory would let another user feed the bar.
  readonly property string runtimeDir: Quickshell.env("XDG_RUNTIME_DIR") || ""
  readonly property string statePath: runtimeDir ? runtimeDir + "/m365bar/state.json" : ""
  readonly property bool tooltipHovered: mouseArea.containsMouse
  readonly property string calendarCommand: setting("calendarCommand",
    "omarchy-launch-or-focus-webapp 'Calendar.*Outlook' 'https://outlook.cloud.microsoft/calendar/view/day'")
  readonly property int maxLabelWidth: Number(setting("maxWidth", 380))
  readonly property bool showWhenEmpty: setting("showWhenEmpty", true) === true

  property var model: null
  property double now: Date.now()
  property int slideIndex: 0

  readonly property var view: Model.compute(model, now, {
    soonMinutes: Number(setting("soonMinutes", 15)),
    urgentMinutes: Number(setting("urgentMinutes", 5)),
    joinLeadMinutes: Number(setting("joinLeadMinutes", 10)),
    lookaheadHours: Number(setting("lookaheadHours", 18))
  })
  readonly property var slide: view.slides.length ? view.slides[slideIndex % view.slides.length] : null

  // What is on screen. Countdown ticks update it in place; a change of meeting
  // (or slide) goes through the crossfade instead.
  property var shown: null
  property string shownKey: ""

  function keyOf(s) { return s ? s.kind + "|" + s.subject : "" }

  function levelColor(level) {
    var fg = bar ? bar.barForeground : Commons.Color.foreground
    if (level === "urgent") return Commons.Color.urgent
    if (level === "accent") return Commons.Color.accent
    return fg
  }

  // Keep an open tooltip current as countdowns tick (fires only on real change).
  readonly property string tooltipText: view.tooltip
  onTooltipTextChanged: if (tooltipHovered && bar) bar.showTooltip(root, tooltipText)
  // React to tooltipHovered itself, not the MouseArea signal: the bar re-checks
  // target.tooltipHovered, and from containsMouseChanged that binding may not
  // have updated yet, so the tooltip was silently dropped.
  onTooltipHoveredChanged: {
    if (!bar) return
    if (tooltipHovered) bar.showTooltip(root, tooltipText)
    else bar.hideTooltip(root)
  }

  onSlideChanged: {
    var key = keyOf(slide)
    if (shown && slide && key !== shownKey && Style.duration(180) > 0) {
      crossfade.restart()
    } else {
      shown = slide
      shownKey = key
    }
  }

  function refresh() {
    Util.execArgv(["systemctl", "--user", "kill", "-s", "USR1", "m365bar.service"])
  }

  function signIn() {
    bar.run("omarchy-launch-floating-terminal-with-presentation m365bar login")
  }

  function join() {
    var url = Model.safeJoin(view.joinUrl)
    if (!url) return
    // Argv, not a shell string: the URL came from a meeting invite.
    // Teams meetings open straight into the web pre-join screen in a Teams
    // web-app window (same browser profile, so already signed in).
    var web = Model.teamsWebJoin(url)
    var teams = web || /^https:\/\/teams\.(microsoft|live)\.com\//i.test(url)
    Util.execArgv([teams ? "omarchy-launch-webapp" : "xdg-open", web || url])
  }

  visible: !vertical && (view.slides.length > 0 || (showWhenEmpty && view.status === "none"))
  implicitWidth: visible ? Math.min(maxLabelWidth, content.implicitWidth) + Style.spacing.controlPaddingX * 2 : 0
  implicitHeight: barSize

  Behavior on implicitWidth {
    NumberAnimation { duration: Style.duration(180); easing.type: Easing.OutCubic }
  }

  FileView {
    id: stateFile
    path: root.statePath
    watchChanges: true
    printErrors: false
    onLoaded: {
      try { root.model = JSON.parse(text()) } catch (e) { /* half-written; next tick */ }
    }
    onLoadFailed: root.model = null
    onFileChanged: reload()
  }

  // Countdown tick. The file watch delivers updates; the slow reload only
  // covers a watch lost when the daemon atomically replaces the file (and the
  // daemon starting after the shell).
  Timer {
    interval: 5000
    running: true
    repeat: true
    onTriggered: root.now = Date.now()
  }

  Timer {
    interval: 30000
    running: root.statePath !== ""
    repeat: true
    onTriggered: stateFile.reload()
  }

  Timer {
    interval: Math.max(2, Number(root.setting("rotateSeconds", 6))) * 1000
    running: root.view.slides.length > 1
    repeat: true
    onTriggered: root.slideIndex = (root.slideIndex + 1) % Math.max(1, root.view.slides.length)
  }

  SequentialAnimation {
    id: crossfade
    NumberAnimation { target: content; property: "opacity"; to: 0; duration: Style.duration(120); easing.type: Easing.OutCubic }
    ScriptAction { script: { root.shown = root.slide; root.shownKey = root.keyOf(root.slide) } }
    NumberAnimation { target: content; property: "opacity"; to: 1; duration: Style.duration(180); easing.type: Easing.OutCubic }
  }

  Row {
    id: content
    anchors.left: parent.left
    anchors.leftMargin: Style.spacing.controlPaddingX
    anchors.verticalCenter: parent.verticalCenter
    spacing: Style.space(6)

    readonly property color tint: root.shown ? root.levelColor(root.shown.level) : (root.bar ? root.bar.barForeground : Commons.Color.foreground)
    readonly property real textOpacity: !root.shown ? 0.45 : root.shown.level === "dim" ? 0.6 : 1

    // Live dot while in a meeting/call, calendar glyph otherwise.
    Item {
      width: glyph.implicitWidth
      height: glyph.implicitHeight
      anchors.verticalCenter: parent.verticalCenter

      Text {
        id: glyph
        textFormat: Text.PlainText
        text: root.shown && root.shown.live ? "●" : (root.view.status === "auth" ? "󰀦" : "󰃭")
        color: content.tint
        opacity: content.textOpacity
        font.family: root.bar ? root.bar.fontFamily : Style.font.family
        font.pixelSize: Style.font.body

        SequentialAnimation on opacity {
          running: root.shown !== null && root.shown.live === true && Style.duration(1) > 0
          loops: Animation.Infinite
          NumberAnimation { to: 0.35; duration: 900; easing.type: Easing.InOutSine }
          NumberAnimation { to: 1; duration: 900; easing.type: Easing.InOutSine }
        }
      }
    }

    Text {
      id: label
      textFormat: Text.PlainText
      anchors.verticalCenter: parent.verticalCenter
      width: Math.min(implicitWidth, root.maxLabelWidth - glyph.implicitWidth - content.spacing)
      text: root.shown ? root.shown.text : ""
      visible: text !== ""
      elide: Text.ElideRight
      color: content.tint
      opacity: content.textOpacity
      font.family: root.bar ? root.bar.fontFamily : Style.font.family
      font.pixelSize: Style.font.body
      font.bold: root.shown !== null && root.shown.level === "urgent"
    }
  }

  MouseArea {
    id: mouseArea
    anchors.fill: parent
    hoverEnabled: true
    acceptedButtons: Qt.LeftButton | Qt.MiddleButton | Qt.RightButton
    cursorShape: Qt.PointingHandCursor

    onClicked: function(mouse) {
      if (!root.bar) return
      if (mouse.button === Qt.MiddleButton) root.refresh()
      else if (mouse.button === Qt.RightButton) root.join()
      else if (root.view.status === "auth") root.signIn()
      else root.bar.run(root.calendarCommand)
    }
  }
}
