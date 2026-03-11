#!/bin/bash
# Navigate to the project directory
cd /home/fet/all-projects/faculty_mis

# Pull the latest changes
git pull origin main

# Activate virtual environment and update dependencies
source venv/bin/activate
pip install -r requirements.txt

# Restart the application with PM2
pm2 restart faculty_mis
