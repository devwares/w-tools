@ECHO OFF

if "%1"=="" goto missing
if "%1"=="/?" goto help
if "%1"=="/h" goto help
if "%1"=="/H" goto help
if "%1"=="1" goto set
if "%1"=="0" goto set
goto wrong

:set
C:\Windows\System32\cmd.exe /k %windir%\System32\reg.exe ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v ConsentPromptBehaviorAdmin /t REG_DWORD /d %1 /f
goto end

:missing
echo Missing parameter (/? for help)
goto end

:wrong
echo Wrong parameter (/? for help)
goto end

:help
echo Sets ConsentPromptBehaviorAdmin policy for local machine
echo.
echo %0 [0-1] [/?]
echo.
echo   1 :  Enables ConsentPromptBehaviorAdmin on local machine
echo   0 :  Disables ConsentPromptBehaviorAdmin on local machine

:end