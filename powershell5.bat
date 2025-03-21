@echo off
cd /d %~dp0
powershell -ExecutionPolicy Bypass -Command ./AllpairsExcelPS5.ps1 %~nx1
pause