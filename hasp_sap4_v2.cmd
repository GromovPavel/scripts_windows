@echo off

IF EXIST "%ProgramFiles%\1cv8\conf\nethasp.ini" (
    copy %~dp0nethasp.ini "%ProgramFiles%\1cv8\conf" /Z /Y
   )
   
IF EXIST "%ProgramFiles(x86)%\1cv8\conf\nethasp.ini" (
    copy %~dp0nethasp.ini "%ProgramFiles(x86)%\1cv8\conf" /Z /Y
   )