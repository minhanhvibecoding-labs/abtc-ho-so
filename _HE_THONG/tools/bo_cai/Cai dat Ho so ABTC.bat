@echo off
rem Cai dat / mo ung dung Ho so ABTC. Bam dup file nay.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_cai_dat\cai_dat.ps1"
if errorlevel 1 pause
