@echo off
rem Run the agent by hand, e.g.:  scripts\run.cmd discover --dry-run
cd /d "%~dp0.."
".venv\Scripts\python.exe" -m origincrest_agent %*
