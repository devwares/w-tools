@ECHO OFF

if "%1"=="" goto missing
if "%1"=="/?" goto help
if "%1"=="/s" goto show
if "%1"=="/S" goto show

echo -- 64 bits -----
%windir%\System32\netsh winhttp show proxy
echo New system proxy configuration :
%windir%\System32\netsh winhttp set proxy %1
echo -- 32 bits -----
%windir%\SysWOW64\netsh.exe winhttp show proxy
echo New system proxy configuration :
%windir%\SysWOW64\netsh.exe winhttp set proxy %1
goto end

:show
echo -- 64 bits -----
%windir%\System32\netsh winhttp show proxy
echo -- 32 bits -----
%windir%\SysWOW64\netsh.exe winhttp show proxy
goto end

:missing
echo Proxy informations missing
goto end

:help
echo Sets system proxy informations for 32 and 64 bits applications
echo.
echo %0 [http://proxyservername:port] [none] [/?]
goto end

:end
