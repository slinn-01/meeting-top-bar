import QtQuick
import Quickshell
import Quickshell.Io
import qs.Commons
import qs.Commons as Commons
import qs.Ui
import "Model.js" as Model

// Latest Teams chat message from someone else, Twitch-chat style:
// a dark pill with a colored sender name and the message. New messages slide
// in; read or old ones dim. Fed by the m365bar daemon.
//
// Left click focuses Teams, right click opens that chat, middle click
// dismisses the message until a newer one arrives.
BarWidget {
  id: root
  moduleName: "steve.teams-chat"

  // No /tmp fallback: a shared directory would let another user feed the bar.
  readonly property string runtimeDir: Quickshell.env("XDG_RUNTIME_DIR") || ""
  readonly property string statePath: runtimeDir ? runtimeDir + "/m365bar/state.json" : ""
  readonly property bool tooltipHovered: mouseArea.containsMouse
  readonly property string teamsCommand: setting("teamsCommand",
    "omarchy-launch-or-focus-webapp 'Microsoft Teams' 'https://teams.microsoft.com/v2/'")
  readonly property int maxLabelWidth: Number(setting("maxWidth", 600))

  property var model: null
  property double now: Date.now()
  property string dismissedKey: ""

  readonly property var view: Model.compute(model, now, {
    hideAfterMinutes: Number(setting("hideAfterMinutes", 0)),
    dimAfterMinutes: Number(setting("dimAfterMinutes", 10)),
    dismissedKey: dismissedKey
  })

  // Message on screen; swapped through the slide-in when a new one lands.
  property var shown: null
  property string shownKey: ""

  // One instance per monitor; middle click dismisses on all of them.
  function dismiss() { dismissedKey = view.key }

  readonly property string tooltipText: view.tooltip
  onTooltipTextChanged: if (tooltipHovered && bar) bar.showTooltip(root, tooltipText)

  onViewChanged: {
    if (view.hidden) {
      shown = null
      shownKey = ""
    } else if (view.key !== shownKey) {
      var animate = shown !== null && Style.duration(180) > 0
      shown = view.msg
      shownKey = view.key
      if (animate) arrive.restart()
    } else {
      shown = view.msg   // unread → read, etc.
    }
  }

  readonly property color nameColor: shown ? Qt.hsla(Model.hue(shown.from), 0.55, 0.68, 1) : Commons.Color.accent

  visible: !vertical && shown !== null
  implicitWidth: visible ? Math.min(maxLabelWidth, pillRow.implicitWidth + Style.space(10) * 2) + Style.space(4) : 0
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
      try { root.model = JSON.parse(text()) } catch (e) { }
    }
    onLoadFailed: root.model = null
    onFileChanged: reload()
  }

  // Age/dim tick; the slow reload only covers a lost file watch.
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

  ParallelAnimation {
    id: arrive
    NumberAnimation { target: pill; property: "opacity"; from: 0; to: 1; duration: Style.duration(220); easing.type: Easing.OutCubic }
    NumberAnimation { target: slide; property: "y"; from: Style.space(8); to: 0; duration: Style.duration(220); easing.type: Easing.Bezier; easing.bezierCurve: Style.easeOut }
  }

  Rectangle {
    id: pill
    anchors.verticalCenter: parent.verticalCenter
    anchors.left: parent.left
    anchors.right: parent.right
    anchors.rightMargin: Style.space(4)
    height: root.barSize - Style.space(6)
    radius: Style.space(3)
    color: Qt.rgba(Commons.Color.background.r, Commons.Color.background.g, Commons.Color.background.b, Number(root.setting("pillOpacity", 0.85)))
    border.width: 1
    border.color: Qt.rgba(1, 1, 1, root.shown && root.shown.unread ? 0.10 : 0.04)
    clip: true

    Item {
      id: slide
      width: parent.width
      height: parent.height

      Row {
        id: pillRow
        anchors.left: parent.left
        anchors.leftMargin: Style.space(10)
        anchors.verticalCenter: parent.verticalCenter
        spacing: Style.space(8)
        opacity: root.view.dim ? 0.6 : 1

        Behavior on opacity {
          NumberAnimation { duration: Style.duration(180); easing.type: Easing.OutCubic }
        }

        Text {
          id: nameText
          textFormat: Text.PlainText
          anchors.verticalCenter: parent.verticalCenter
          text: root.shown ? root.shown.from.split(" ")[0] : ""
          color: root.nameColor
          font.family: root.bar ? root.bar.fontFamily : Style.font.family
          font.pixelSize: Style.font.body
          font.bold: true
        }

        Text {
          textFormat: Text.PlainText
          anchors.verticalCenter: parent.verticalCenter
          width: Math.min(implicitWidth, root.maxLabelWidth - nameText.implicitWidth - pillRow.spacing - Style.space(10) * 2 - Style.space(4))
          text: root.shown ? root.shown.text : ""
          elide: Text.ElideRight
          color: root.bar ? root.bar.barForeground : Commons.Color.foreground
          font.family: root.bar ? root.bar.fontFamily : Style.font.family
          font.pixelSize: Style.font.body
        }
      }
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
      if (mouse.button === Qt.MiddleButton) {
        root.broadcast("dismiss")
      } else if (mouse.button === Qt.RightButton) {
        // Argv, not a shell string: the URL came from Graph.
        var url = root.shown ? String(root.shown.webUrl || "") : ""
        if (/^https:\/\/teams\.microsoft\.com\/[^\s'"`\\<>]*$/.test(url)) Util.execArgv(["omarchy-launch-webapp", url])
        else root.bar.run(root.teamsCommand)
      } else {
        root.bar.run(root.teamsCommand)
      }
    }
    // The bar only shows a tooltip while target.tooltipHovered is true.
    onContainsMouseChanged: {
      if (!root.bar) return
      if (containsMouse) root.bar.showTooltip(root, root.view.tooltip)
      else root.bar.hideTooltip(root)
    }
  }
}
