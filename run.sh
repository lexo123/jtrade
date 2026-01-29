#!/bin/bash
# Activate virtual environment and run CLI
cd "$(dirname "$0")"
source venv/bin/activate
python cli.py
