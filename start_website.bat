@echo off
title Naue Geosynthetics - Project Opportunity Tracker
cd /d "%~dp0"
echo Starting Naue Geosynthetics Opportunity Intelligence Portal...
python server.py --open --port 8080
pause

