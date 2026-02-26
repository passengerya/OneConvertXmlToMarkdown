@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0OneConvertPipeline-GUI.ps1"
