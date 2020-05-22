@echo off
REM cd %~dp0
REM REG ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 0 /f
PowerShell -ExecutionPolicy Bypass -File "%~dp0scriptname.ps1"
echo Terminated.
pause