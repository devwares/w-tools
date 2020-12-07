@echo off
SET SCRIPTHOME=%~dp0
SET SCRIPTNAME=script.ps1
SET SCRIPTPATH=%SCRIPTHOME%%SCRIPTNAME%

REM Simple Call 1
PowerShell.exe -ExecutionPolicy Unrestricted -File %SCRIPTPATH%

REM Simple Call 2
PowerShell.exe -ExecutionPolicy Bypass -File "%SCRIPTPATH%"

REM Call Script with example parameter "-ConfigFile"
Powershell.exe -ExecutionPolicy Bypass -Command "& '%SCRIPTPATH%' -ConfigFile '%CONFIGFILE%'"