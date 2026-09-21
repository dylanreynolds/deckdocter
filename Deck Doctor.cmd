@echo off
cd /d "%~dp0"
start "" pwsh -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0Deck-Doctor-GUI.ps1"
