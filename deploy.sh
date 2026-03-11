#!/bin/bash
cd "$(dirname "$0")"
git pull origin main
source venv/bin/activate
pip install -r requirements.txt
pm2 restart faculty_mis