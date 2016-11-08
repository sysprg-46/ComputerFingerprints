@echo off
REM --------------------------------------------------
REM Batch file to run COMPInfo.js, RAMInfo.js
REM --------------------------------------------------
REM %0 it always full path to the batch file
REM but current path can differ those like:
REM e:\Delme\!Projects\CompInfo>k:\_RAM_INFO_\test.cmd
REM --------------------------------------------------
setlocal enableextensions disabledelayedexpansion

REM Save current path in variable StartPath
for /f %%a in ('"cd"') do (set StartPath=%%a)

REM Parse %0 to set variables Path and File
call :ParseCommand %0
set "JSName=%Path%%File%"

REM Set default log file directory and name
set "log=%temp%\%USERDOMAIN%-%File%Info.log"

REM Redefine log file for Well-Known computers
if "%USERDOMAIN%"=="HP-OK" set "log=%JSName%-HP_103C_5335KV_Pavilion_tx2500.log"
if "%USERDOMAIN%"=="R500D" set "log=%JSName%-LENOVO_ThinkPad_R500_2732W12.log"
if "%USERNAME%"=="marin"   set "log=%JSName%-LENOVO_ThinkPad_R61_8932GMG.log"
if "%USERNAME%"=="ninak"   set "log=%JSName%-LENOVO_IDEAPAD_U310.log"
if "%USERNAME%"=="liza"    set "log=%JSName%-DELL-Venue-11PROe.log"

REM Delete old log file if it exist
del %log% 2>nul

REM Change dir to scripts home dir if necessary
if NOT %Path%==%StartPath% cd /d %Path%

REM Run JS-script
echo Running %windir%\System32\CScript.exe %JSName%.js %1 %2 //NOLOGO
%windir%\System32\CScript.exe %JSName%.js %1 %2 //NOLOGO>%log%

if exist %log% Echo log file %log% was created

REM Show log created in the Notepad
cmd.exe /c notepad.exe %log% 

REM Change dir to scripts home dir if necessary
if NOT %Path%==%StartPath% cd /d %StartPath%

goto :EOF

:ParseCommand
set "Path=%~d1%~p1" 
set "File=%~n1"
set "Drive=%~d1"
exit /b