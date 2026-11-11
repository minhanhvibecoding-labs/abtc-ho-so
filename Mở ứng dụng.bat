@echo off
rem Mo ung dung Ho so ABTC (giao dien web chay tren may, khong can cai them).
rem Cua so may chu tu thu nho xuong taskbar; giu no mo trong luc dung. Tat bang nut "Tat ung dung" trong app hoac dong cua so do.
cd /d "%~dp0"
start "Ung dung Ho so ABTC - giu cua so nay mo trong luc dung" /min powershell -NoProfile -ExecutionPolicy Bypass -Command "& '.\_HE_THONG\app\server.ps1'; Write-Host ''; pause"
