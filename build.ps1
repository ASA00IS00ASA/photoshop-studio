# Photoshop Studio build script -> single-file exe (double-click to run, no browser)
# Usage:  powershell -ExecutionPolicy Bypass -File .\build.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "[1/3] generate icon"
python tools/make_icon.py

Write-Host "[2/3] clean previous build"
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
Remove-Item -Force PhotoshopStudio.spec -ErrorAction SilentlyContinue

Write-Host "[3/3] PyInstaller onefile + windowed"
python -m PyInstaller `
  --noconfirm --clean --onefile --windowed `
  --name PhotoshopStudio `
  --icon assets/studio.ico `
  --add-data "gui/web;gui/web" `
  --collect-all webview `
  --hidden-import win32com `
  --hidden-import win32com.client `
  --hidden-import pythoncom `
  --hidden-import pywintypes `
  --hidden-import win32api `
  gui_app.py

$exe = Join-Path $PSScriptRoot "dist\PhotoshopStudio.exe"
if (-not (Test-Path $exe)) { throw "build failed: $exe not found" }

# Files created inside this workspace inherit a Low integrity label, which makes
# PyInstaller's onefile self-extraction fail; promote the exe to Medium.
icacls $exe /setintegritylevel Medium | Out-Null

$size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host ""
Write-Host "OK -> $exe ($size MB)"
Write-Host "Double-click to run. Copy this single exe anywhere to distribute."
