@echo off
echo Snipper Pro - Requesting Administrator privileges...
PowerShell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0register_snipper_pro_simple.ps1\"' -Verb RunAs"
pause 