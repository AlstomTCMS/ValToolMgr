@echo off
cls

rem ************************************************************
rem Installation automatique de la macro au démarrage de session
rem ************************************************************
rem UpdateMacroOnStartup.exe
rem Auteur du fichier: DLE
rem Societe : Alten

set version=v2.0.1
echo Version : %version% 22.04.2013

set updateMacroFileName=C:\macros_alstom\UpdateMacroTCMS.exe

if exist %updateMacroFileName% goto CallUpdate
exit

:CallUpdate
echo call %updateMacroFileName% onStartup %version%
::pause
call %updateMacroFileName% onStartup %version%