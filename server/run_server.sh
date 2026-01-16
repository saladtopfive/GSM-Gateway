#!/bin/bash
set -e

WATCHDOG_INTERVAL=10

# poinformuj systemd że start OK
systemd-notify READY=1

# uruchom Flask w tle
/usr/bin/python3 app.py &
APP_PID=$!

# pętla watchdoga
while kill -0 "$APP_PID" 2>/dev/null; do
    systemd-notify WATCHDOG=1
    sleep $WATCHDOG_INTERVAL
done

# jeśli Python padł – wyjdź (systemd zrestartuje)
exit 1
