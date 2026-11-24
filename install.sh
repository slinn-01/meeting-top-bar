#!/bin/bash
# Install the m365bar daemon + the two bar widgets for the current user.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
plugins="$HOME/.config/omarchy/plugins"

mkdir -p "$HOME/.local/bin" "$HOME/.config/systemd/user" "$HOME/.config/m365bar" "$plugins"
ln -sfn "$here/daemon/m365bar.py" "$HOME/.local/bin/m365bar"
cp "$here/daemon/m365bar.service" "$HOME/.config/systemd/user/m365bar.service"
[[ -f $HOME/.config/m365bar/config.toml ]] || cp "$here/daemon/config.example.toml" "$HOME/.config/m365bar/config.toml"

# Copied, not symlinked. The shell notices the change but can keep running the
# old widget code, so updates need 'omarchy restart shell' (see the echo below).
for p in steve.meetings steve.teams-chat; do
  mkdir -p "$plugins/$p"
  cp "$here/plugins/$p/"{manifest.json,BarWidget.qml,Model.js} "$plugins/$p/"
done

systemctl --user daemon-reload
echo "Installed. Next: set client_id in ~/.config/m365bar/config.toml, run 'm365bar login',"
echo "then: systemctl --user enable --now m365bar.service"
echo "Updating an existing install? Run 'omarchy restart shell' to load the new widget code."
