#!/bin/bash

# Script to start the Flask app and Cloudflare Tunnel together
# This version monitors for the public URL and displays it

set -e  # Exit on any error

echo "Starting Excel Template Generator with public access..."

# Clean up any previous log files
rm -f tunnel_full.log

# Define the port to use
PORT=5000  # Use port 5000 as requested

# Start the Flask app in the background
echo "Starting Flask app on port $PORT..."
source venv/bin/activate
python -c "
import os
from app import app
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=$PORT)
" > web.log 2>&1 &
FLASK_PID=$!

# Wait a moment for the Flask app to start
sleep 3

# Check if the Flask app started successfully
if ! kill -0 $FLASK_PID 2>/dev/null; then
    echo "Error: Flask app failed to start. Check web.log for details."
    exit 1
fi

echo "Flask app started with PID: $FLASK_PID on port $PORT"

# Start the tunnel in the background
echo "Starting Cloudflare Tunnel..."
if [ ! -f "./cloudflared" ]; then
    echo "Error: cloudflared not found. Please run setup_tunnel.sh first."
    exit 1
fi

# Function to extract URL from log file
extract_url() {
    # Look for various possible URL patterns in the log
    grep -oE 'https://[a-z0-9-]+\.trycloudflare\.com' tunnel_full.log 2>/dev/null | head -n 1
}

# Run the tunnel in the background with combined stdout/stderr log
stdbuf -oL -eL ./cloudflared tunnel --url http://localhost:$PORT > tunnel_full.log 2>&1 &
TUNNEL_PID=$!

echo "Cloudflare Tunnel started with PID: $TUNNEL_PID"

# Save PIDs to a file for easy stopping
echo "$FLASK_PID $TUNNEL_PID $PORT" > web.pid

# Monitor the log file for the public URL
echo ""
echo "Waiting for public URL..."
echo "This may take 10-30 seconds..."
echo ""

# Wait up to 60 seconds for the URL to appear
COUNT=0
MAX_WAIT=30  # 30 iterations with 2-second intervals
FOUND_URL=""

while [ $COUNT -lt $MAX_WAIT ]; do
    FOUND_URL=$(extract_url)
    if [ -n "$FOUND_URL" ]; then
        break
    fi
    sleep 2
    COUNT=$((COUNT + 1))
done

if [ -n "$FOUND_URL" ]; then
    echo "Public URL is ready:"
    echo ""
    echo "  $FOUND_URL"
    echo ""
    echo "You can now access the app from your phone using this URL."
    echo ""
else
    echo "Could not automatically detect the public URL."
    echo "Please check the tunnel log manually:"
    echo "  cat tunnel_full.log | grep -i cloudflare"
    echo ""
    # Show the most recent lines that might contain the URL
    tail -20 tunnel_full.log | grep -i cloudflare || echo "No cloudflare URLs found in recent logs"
fi

# Keep the script running to maintain the processes
echo "Both applications are running. Press Ctrl+C to stop."
echo ""
if [ -n "$FOUND_URL" ]; then
    echo "Local access: http://localhost:$PORT"
    echo "Public access: $FOUND_URL"
else
    echo "Local access: http://localhost:$PORT"
    echo "Public access: Check tunnel_full.log for URL"
fi
echo ""

# Wait for interruption signal
trap 'echo -e "\nStopping applications..."; kill $FLASK_PID $TUNNEL_PID 2>/dev/null; rm -f tunnel_full.log; exit' INT TERM

# Wait indefinitely
while kill -0 $FLASK_PID 2>/dev/null && kill -0 $TUNNEL_PID 2>/dev/null; do
    sleep 5
done