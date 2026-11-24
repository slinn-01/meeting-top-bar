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

  // Hover opens our own dropdown under the pill (the bar's tooltip is
  // centred plain text as wide as the longest message). A short delay keeps
  // a passing pointer from flashing it.
  readonly property bool hovered: mouseArea.containsMouse
  property bool dropdownOpen: false
  onHoveredChanged: {
    if (hovered) openDelay.restart()
    else { openDelay.stop(); dropdownOpen = false }
  }
  Timer {
    id: openDelay
    interval: 300
    onTriggered: root.dropdownOpen = root.hovered
  }

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
  }

  // Recent messages, right-aligned under the pill and at least dropdownWidth
  // wide. Passive: it never takes input, so it can't steal bar clicks.
  PopupWindow {
    id: dropdown

    readonly property var barWindow: root.QsWindow.window
    readonly property int pad: Style.space(12)
    readonly property color text: Commons.Color.tooltip.text
    readonly property string fontFamily: root.bar ? root.bar.fontFamily : Style.font.family

    visible: root.dropdownOpen && root.visible && !!barWindow
    color: "transparent"
    mask: Region {}
    implicitWidth: Math.round(Math.max(pill.width, Number(root.setting("dropdownWidth", 440))))
    implicitHeight: Math.ceil(card.implicitHeight)

    anchor {
      id: dropdownAnchor
      window: dropdown.barWindow
      adjustment: PopupAdjustment.Slide
      edges: Edges.Top | Edges.Left
      gravity: Edges.Bottom | Edges.Right
      rect.width: 1
      rect.height: 1

      onAnchoring: {
        if (!dropdown.barWindow) return
        var gap = Style.space(6)
        var y = root.bar && root.bar.position === "bottom" ? -dropdown.implicitHeight - gap : pill.height + gap
        var p = dropdown.barWindow.contentItem.mapFromItem(pill, pill.width - dropdown.implicitWidth, y)
        dropdownAnchor.rect.x = Math.round(p.x)
        dropdownAnchor.rect.y = Math.round(p.y)
      }
    }

    // As the bar's own tooltip does: a fresh popup surface first maps at a
    // whole-number scale, so stay transparent until it matches the bar's.
    readonly property bool scaleReady: visible && Math.round(devicePixelRatio * 120) === Math.round(barWindow.devicePixelRatio * 120)
    property bool revealed: false
    onVisibleChanged: if (!visible) revealed = false
    onScaleReadyChanged: if (scaleReady) Qt.callLater(function() {
      if (dropdown.scaleReady) dropdown.revealed = true
    })

    Rectangle {
      id: card
      width: dropdown.implicitWidth
      implicitHeight: list.implicitHeight + dropdown.pad * 2
      opacity: dropdown.revealed ? 1 : 0
      color: Commons.Color.tooltip.background
      border.width: 1
      border.color: Qt.rgba(Commons.Color.tooltip.border.r, Commons.Color.tooltip.border.g, Commons.Color.tooltip.border.b, 0.35)
      radius: Style.cornerRadius

      Behavior on opacity {
        NumberAnimation { duration: Style.duration(120); easing.type: Easing.OutCubic }
      }

      Column {
        id: list
        x: dropdown.pad
        y: dropdown.pad
        width: parent.width - dropdown.pad * 2
        spacing: Style.space(12)

        Repeater {
          model: root.view.entries

          delegate: Column {
            id: entry
            required property var modelData
            width: list.width
            spacing: Style.space(3)
            opacity: modelData.unread ? 1 : 0.7

            // Header: ● Name · topic ........ 5m
            Item {
              width: parent.width
              height: nameLabel.implicitHeight

              Rectangle {
                id: unreadDot
                visible: entry.modelData.unread
                width: Style.space(6)
                height: width
                radius: width / 2
                color: Commons.Color.accent
                anchors.verticalCenter: parent.verticalCenter
              }

              Text {
                id: nameLabel
                x: unreadDot.visible ? unreadDot.width + Style.space(6) : 0
                textFormat: Text.PlainText
                text: entry.modelData.from
                color: Qt.hsla(entry.modelData.hue, 0.55, 0.68, 1)
                font.family: dropdown.fontFamily
                font.pixelSize: Style.font.body
                font.bold: true
              }

              Text {
                anchors.left: nameLabel.right
                anchors.leftMargin: Style.space(6)
                anchors.right: agoLabel.left
                anchors.rightMargin: Style.space(8)
                anchors.baseline: nameLabel.baseline
                textFormat: Text.PlainText
                text: entry.modelData.topic ? "in " + entry.modelData.topic : ""
                elide: Text.ElideRight
                color: dropdown.text
                opacity: 0.55
                font.family: dropdown.fontFamily
                font.pixelSize: Style.font.bodySmall
              }

              Text {
                id: agoLabel
                anchors.right: parent.right
                anchors.baseline: nameLabel.baseline
                textFormat: Text.PlainText
                text: entry.modelData.ago
                color: dropdown.text
                opacity: 0.55
                font.family: dropdown.fontFamily
                font.pixelSize: Style.font.bodySmall
              }
            }

            Text {
              x: nameLabel.x
              width: parent.width - x
              textFormat: Text.PlainText
              text: entry.modelData.text
              wrapMode: Text.Wrap
              maximumLineCount: 3
              elide: Text.ElideRight
              lineHeight: 1.15
              color: dropdown.text
              font.family: dropdown.fontFamily
              font.pixelSize: Style.font.body
            }
          }
        }

        Text {
          visible: root.view.entries.length === 0
          textFormat: Text.PlainText
          text: "No recent Teams messages"
          color: dropdown.text
          opacity: 0.6
          font.family: dropdown.fontFamily
          font.pixelSize: Style.font.body
        }

        Repeater {
          model: root.view.notes
          delegate: Text {
            required property var modelData
            width: list.width
            textFormat: Text.PlainText
            text: "⚠ " + modelData
            wrapMode: Text.Wrap
            color: Commons.Color.urgent
            font.family: dropdown.fontFamily
            font.pixelSize: Style.font.bodySmall
          }
        }

        Rectangle {
          width: list.width
          height: 1
          color: dropdown.text
          opacity: 0.12
        }

        Text {
          width: list.width
          textFormat: Text.PlainText
          text: "Click: Teams   ·   Right-click: open chat   ·   Middle-click: dismiss"
          elide: Text.ElideRight
          horizontalAlignment: Text.AlignHCenter
          color: dropdown.text
          opacity: 0.45
          font.family: dropdown.fontFamily
          font.pixelSize: Style.font.caption
        }
      }
    }
  }
}
