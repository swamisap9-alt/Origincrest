# One-time setup on Windows. Run from the repo folder:
#   powershell -ExecutionPolicy Bypass -File scripts\setup.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (Get-Command py -ErrorAction SilentlyContinue) { $pyExe = "py"; $pyArgs = @("-3") }
else { $pyExe = "python"; $pyArgs = @() }
$version = & $pyExe @pyArgs -c "import sys; print('%d.%d' % sys.version_info[:2])"
if ([version]$version -lt [version]"3.10") {
    throw "Python 3.10 or newer is needed (found $version). Install it from https://www.python.org/downloads/"
}

if (-not (Test-Path ".venv")) {
    & $pyExe @pyArgs -m venv .venv
}
& .\.venv\Scripts\python.exe -m pip install --upgrade pip --quiet
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt --quiet

if (-not (Test-Path "config.yaml")) {
    Copy-Item "config.example.yaml" "config.yaml"
    Write-Host "Created config.yaml - fill in the <<EDIT ME>> values."
}
if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "Created .env - add your Anthropic API key and IMAP app password."
}
Write-Host ""
Write-Host "Setup done. Next:"
Write-Host "  1. Edit config.yaml and .env"
Write-Host "  2. scripts\run.cmd check --api"
Write-Host "  3. scripts\run.cmd discover --dry-run"
