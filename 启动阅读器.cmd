@echo off
chcp 65001 >nul
start "FishReader" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0FishReader.ps1"
