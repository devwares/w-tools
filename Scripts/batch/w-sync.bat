@ECHO OFF

if "%1"=="" goto missing
if "%1"=="/?" goto help

if not exist %1 goto unknown

REM *****************************************************
REM Make environment variable changes local to this batch file
REM *****************************************************
SETLOCAL

REM *****************************************************
REM Remember original dir
REM *****************************************************
SET SYNCBASEDIR=%CD%

REM *****************************************************
REM CSV CONFIG FILE FORMAT
REM
REM EMPTY VALUES MUST ABSOLUTLY BE FILLED WITH SPACES !
REM Local drive letters must be preceded by "/cygdrive/" and followed by slash ("/")
REM Antislashes ("\") must be replaces by slashes ("/")
REM e.g "C:\Temp" becomes "/cygdrive/C/Temp"
REM
REM 1 : sync type (1 = local-ssh, 2 = ssh-local, 3 = local-local)
REM 2 : source directory/path
REM 3 : target directory/path
REM 4 : ssh server (ssh sync only)
REM 5 : port ssh (ssh sync only)
REM 6 : user ssh (ssh sync only)
REM 7 : fichier clef rsa (ssh sync only)
REM 8 : additional rsync arguments (optional)
REM
REM Example configuration lines :
REM 1,"/cygdrive/C/users/john/Documents","/home/john/backup",server.domain.com,22,john,"/cygdrive/C/Users/john/Documents/clefs/id_rsa",--dry-run
REM 2,"/home/john/backup/docs/","/cygdrive/C/users/john/Documents",server.domain.com,22,john,"/cygdrive/C/Users/john/Documents/clefs/id_rsa",--dry-run
REM 3,"/cygdrive/C/users/john/test-archive","/cygdrive/D/test-backup", , , , ,--dry-run
REM *****************************************************
set SYNC_LIST=%1

REM *****************************************************
REM Optional sh script file path to be executed on remote host after ssh sync
REM *****************************************************
set SYNC_SHSCRIPT=%2

REM *****************************************************
REM Specify where to endd rsync and related files
REM Default value is the directory of this batch file
REM *****************************************************
SET SCRIPT_HOME=%~dp0
SET CWOLDPATH=%PATH%
CD /D %SCRIPT_HOME%..\..\Development tools\cygwin
SET CWRSYNCHOME=%CD%

REM *****************************************************
REM Create a home directory for .ssh 
REM *****************************************************
IF NOT EXIST "%CWRSYNCHOME%\home\%USERNAME%" MKDIR "%CWRSYNCHOME%\home\%USERNAME%"
IF NOT EXIST "%CWRSYNCHOME%\home\%USERNAME%\.ssh" MKDIR "%CWRSYNCHOME%\home\%USERNAME%\.ssh"

REM *****************************************************
REM Make cygwin home as a part of system PATH to find required DLLs
REM *****************************************************
SET PATH="%CWRSYNCHOME%\bin";%PATH%

REM *****************************************************
REM Windows paths may contain a colon (:) as a part of drive designation and 
REM backslashes (example c:\, g:\). However, in rsync syntax, a colon in a 
REM path means searching for a remote host. Solution: use absolute path 'a la unix', 
REM replace backslashes (\) with slashes (/) and put -- in front of the 
REM drive letter:
REM 
REM Example : C:\WORK\* --> c/work/*
REM 
REM Example 1 - rsync recursively to a unix server with an openssh server :
REM
REM       rsync -r c/work/ remotehost:/home/user/work/
REM
REM Example 2 - Local rsync recursively 
REM
REM       rsync -r c/work/ d/work/doc/
REM
REM Example 3 - rsync to an rsync server recursively :
REM    (Double colons?? YES!!)
REM
REM       rsync -r c/doc/ remotehost::module/doc
REM
REM Rsync is a very powerful tool. Please look at documentation for other options. 
REM
REM *****************************************************

REM *****************************************************
REM START PROCESSING
REM *****************************************************
if [%SYNC_SHSCRIPT%] NEQ [] (
	if "%SYNC_SHSCRIPT%"=="" goto missing
	echo Script file %SYNC_SHSCRIPT% will be executed after transfer is done
)

echo.
echo _______________________________________________________________________________
echo DEBUT DU TRAITEMENT %date% A %time:~0,5%

FOR /f "tokens=1,2,3,4,5,6,7,8 delims=," %%a in (%SYNC_LIST%) do (

	if %%a EQU 1 (
	
		echo.
		echo ############################ RSYNC : LOCAL TO SSH #############################
		echo #                                                                             #
		echo # %%b --^> %%d:%%c
		echo #                                                                             #
		echo ###############################################################################
		echo.
		rsync %%h -e 'ssh -i %%g -p %%e' -r -t -v --progress --delete -s %%b %%f@%%d:%%c
		echo.
		REM if not "%SYNC_SHSCRIPT%"==" " ssh %%f@%%d -i %%g -p%%e "sh -s " < %SYNC_SHSCRIPT%
	)
	
	if %%a EQU 2 (

		echo.
		echo ############################ RSYNC : SSH TO LOCAL #############################
		echo #                                                                             #
		echo # %%d:%%b --^> %%c
		echo #                                                                             #
		echo ###############################################################################
		echo.
		rsync %%h -e 'ssh -i %%g -p %%e' -r -t -v --progress --delete -s %%f@%%d:%%b %%c
		echo.
		REM if not "%SYNC_SHSCRIPT%"==" " ssh %%f@%%d -i %%g -p%%e "sh -s " < %SYNC_SHSCRIPT%
		
	)
	
	if %%a EQU 3 (
		
		echo.
		echo ########################### RSYNC : LOCAL TO LOCAL ############################
		echo #                                                                             #
		echo # %%b --^> %%c
		echo #                                                                             #
		echo ###############################################################################
		echo.
		rsync %%h -r -t -v --progress --delete -s "%%b" "%%c"
	
	)

)

echo.
echo.
echo SYNC STOPPED %date% AT %time:~0,5%
echo _______________________________________________________________________________

REM RESTORE INITIAL VALUES
SET PATH=%CWOLDPATH%
CD /D %SYNCBASEDIR%
goto end

:help
echo Reads CSV file for rsync configuration(s) to process file transfers
echo.
echo %0 config_file [sh_script_file] [/?]
goto end

:missing
echo CSV config file path missing
goto end

:unknown
echo File not found - %1
goto end

:end
