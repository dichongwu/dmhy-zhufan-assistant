#!/bin/zsh

PLIST="$HOME/Library/LaunchAgents/com.dmhy.zhufan.plist"
LABEL="com.dmhy.zhufan"
URL="http://127.0.0.1:8766"

if ! launchctl print "gui/$(id -u)/$LABEL" >/dev/null 2>&1; then
  launchctl bootstrap "gui/$(id -u)" "$PLIST"
fi

sleep 1
open "$URL"
