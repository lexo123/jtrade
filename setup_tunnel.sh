#!/bin/bash

# Script to download and set up Cloudflare Tunnel for public access

set -e  # Exit on any error

echo "Setting up Cloudflare Tunnel..."

# Detect the operating system
OS_TYPE=$(uname -s | tr '[:upper:]' '[:lower:]')

# Detect the architecture
ARCH=$(uname -m)
case $ARCH in
    x86_64)
        ARCH="amd64"
        ;;
    aarch64|armv8l)
        ARCH="arm64"
        ;;
    armv7l)
        ARCH="arm"
        ;;
    *)
        echo "Unsupported architecture: $ARCH"
        exit 1
        ;;
esac

# Download the appropriate version of cloudflared
DOWNLOAD_URL="https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-${ARCH}"

echo "Downloading Cloudflare Tunnel (${OS_TYPE}-${ARCH})..."
curl -L --output cloudflared ${DOWNLOAD_URL}

# Make it executable
chmod +x cloudflared

echo "Cloudflare Tunnel downloaded and installed successfully!"

echo ""
echo "To run the app with public access:"
echo "1. Start your Flask app: python3 app.py"
echo "2. In another terminal, run: ./cloudflared tunnel --url http://localhost:5000"
echo ""
echo "The tunnel command will provide a public URL that you can access from anywhere!"