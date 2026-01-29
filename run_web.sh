#!/bin/bash
# Start the Flask web app
cd "$(dirname "$0")"
source venv/bin/activate
python app.py
