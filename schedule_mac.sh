#!/bin/bash
# Installs (or removes) a daily run of jobs.py on macOS using launchd.
#   ./schedule_mac.sh install 09:00     -> run every day at 09:00
#   ./schedule_mac.sh remove
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
LABEL="com.local.jobfinder"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"

case "$1" in
  install)
    TIME="${2:-09:00}"; H="${TIME%%:*}"; M="${TIME##*:}"
    mkdir -p "$HOME/Library/LaunchAgents"
    cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key><array>
    <string>$(command -v python3)</string><string>$DIR/jobs.py</string>
  </array>
  <key>WorkingDirectory</key><string>$DIR</string>
  <key>StartCalendarInterval</key><dict>
    <key>Hour</key><integer>$((10#$H))</integer><key>Minute</key><integer>$((10#$M))</integer>
  </dict>
  <key>StandardOutPath</key><string>$DIR/output/run.log</string>
  <key>StandardErrorPath</key><string>$DIR/output/run.log</string>
</dict></plist>
EOF
    mkdir -p "$DIR/output"
    launchctl unload "$PLIST" 2>/dev/null || true
    launchctl load "$PLIST"
    echo "Scheduled daily at $TIME. Log: $DIR/output/run.log"
    echo "Note: the Mac must be awake at that time (launchd runs missed jobs on wake)."
    ;;
  remove)
    launchctl unload "$PLIST" 2>/dev/null || true
    rm -f "$PLIST"
    echo "Schedule removed."
    ;;
  *) echo "Usage: $0 install [HH:MM] | remove"; exit 1;;
esac
