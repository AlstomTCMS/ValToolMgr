@echo off
cls

rem ************************************************************
rem Installation automatique de la macro au démarrage de session
rem ************************************************************
rem UpdateMacroOnStartup.exe
rem Auteur du fichier: DLE
rem Societe : Alten

set version=A8
echo Version : %version% 12.02.2013

set updateMacroFileName=C:\macros_alstom\UpdateMacroTCMS.exe

if exist %updateMacroFileName% goto CallUpdate
exit

:CallUpdate
echo call %updateMacroFileName% onStartup %version%
::pause
call %updateMacroFileName% onStartup %version%