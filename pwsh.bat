@echo off
cd /d %~dp0
pwsh ./AllpairsExcelPS7.ps1 %~nx1
pause