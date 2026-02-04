#!/bin/bash

# Script to stop the Flask app and Cloudflare Tunnel

echo "Stopping Excel Template Generator..."

# Kill all child processes
pkill -P $$

# Also try to kill any remaining processes from the PID file
if [ -f "web.pid" ]; then
    PIDS=$(cat web.pid)
    for pid in $PIDS; do
        kill $pid 2>/dev/null || true
    done
    rm web.pid
fi

# Kill any remaining cloudflared processes
pkill cloudflared 2>/dev/null || true

echo "Application stopped."