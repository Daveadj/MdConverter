@echo off
REM Serves this folder over http:// rather than opening index.html directly.
REM pdf.js loads its worker from a cross-origin CDN, which file:// blocks.
setlocal
cd /d "%~dp0"

set PORT=8770
where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found on PATH.
  echo Alternatives: "npx serve -l %PORT%" or any static web server in this folder.
  pause
  exit /b 1
)

echo Markdown Converter  -  http://localhost:%PORT%/
echo Press Ctrl+C to stop.
start "" http://localhost:%PORT%/
python -m http.server %PORT% --bind 127.0.0.1
