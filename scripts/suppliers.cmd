@echo off
rem Run the supplier agent by hand, e.g.:  scripts\suppliers.cmd search --dry-run
cd /d "%~dp0.."
".venv\Scripts\python.exe" -m origincrest_agent.supplier_agent %*
