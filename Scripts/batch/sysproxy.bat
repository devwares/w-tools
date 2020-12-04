@ECHO OFF

if "%1"=="" goto missing
if "%1"=="/?" goto help
if "%1"=="/r" goto reset
if "%1"=="/R" goto reset
if "%1"=="/s" goto show
if "%1"=="/S" goto show

:set
echo -- 64 bits -----
%windir%\System32\netsh winhttp set proxy %1 %2
echo -- 32 bits -----
%windir%\SysWOW64\netsh.exe winhttp set proxy %1 %2
goto end

:reset
echo -- 64 bits -----
%windir%\System32\netsh winhttp reset proxy
echo -- 32 bits -----
%windir%\SysWOW64\netsh.exe winhttp reset proxy
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
echo %0 [proxy informations] [bypass-list] [/R] [/S] [/?]
echo.
echo   [proxy informations] : 	Required.
echo	  				Specifies the proxy server to use for http,
echo	  				secure http (https), or both http and https protocols
echo.
echo   [bypass-list] :	Optional.
echo  			Specifies a list of Web sites that should be visited
echo  			without utilizing the proxy server. Use "<local>" to
echo  			bypass all short name hosts.
echo.
echo   /R : reset WinHTTP proxy to DIRECT
echo.
echo   /S : shows WinHTTP proxy current configuration
echo.
echo   Examples :	%0 myproxy
echo  			%0 myproxy:80 "<local>bar"
echo  			%0 "http=myproxy;https=sproxy:88" "*.contoso.com"
goto end

:end
