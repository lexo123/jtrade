#!/bin/bash

# Script to start the Flask app and Cloudflare Tunnel together

set -e  # Exit on any error

echo "Starting Excel Template Generator with public access..."

# Start the Flask app in the background
echo "Starting Flask app on port 5000..."
source venv/bin/activate
python app.py > web.log 2>&1 &
FLASK_PID=$!

# Wait a moment for the Flask app to start
sleep 3

# Check if the Flask app started successfully
if ! kill -0 $FLASK_PID 2>/dev/null; then
    echo "Error: Flask app failed to start. Check web.log for details."
    exit 1
fi

echo "Flask app started with PID: $FLASK_PID"

# Start the tunnel
echo "Starting Cloudflare Tunnel..."
if [ ! -f "./cloudflared" ]; then
    echo "Error: cloudflared not found. Please run setup_tunnel.sh first."
    exit 1
fi

# Run the tunnel in the background
./cloudflared tunnel --url http://localhost:5000 > tunnel.log 2>&1 &
TUNNEL_PID=$!

echo "Cloudflare Tunnel started with PID: $TUNNEL_PID"

# Save PIDs to a file for easy stopping
echo "$FLASK_PID $TUNNEL_PID" > web.pid

echo ""
echo "Application is now running!"
echo "- Local access: http://localhost:5000"
echo "- Public access: Will be shown in tunnel logs (check tunnel.log)"
echo ""
echo "To stop the application, run: pkill -P $$"
echo ""

# Wait for both processes
wait $FLASK_PID $TUNNEL_PID