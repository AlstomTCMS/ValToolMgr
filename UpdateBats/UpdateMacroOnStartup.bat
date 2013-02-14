@ECHO OFF
REM BFCPEOPTIONSTART
REM Advanced BAT to EXE Converter www.BatToExeConverter.com
REM BFCPEEXE=
REM BFCPEICON=
REM BFCPEICONINDEX=0
REM BFCPEEMBEDDISPLAY=0
REM BFCPEEMBEDDELETE=1
REM BFCPEVERINCLUDE=0
REM BFCPEVERVERSION=1.0.0.0
REM BFCPEVERPRODUCT=Product Name
REM BFCPEVERDESC=Product Description
REM BFCPEVERCOMPANY=Your Company
REM BFCPEVERCOPYRIGHT=Copyright Info
REM BFCPEOPTIONEND
@ECHO ON
@echo off
cls

rem ************************************************************
rem Installation automatique de la macro au démarrage de session
rem ************************************************************
rem UpdateMacroOnStartup.exe
rem Auteur du fichier: DLE
rem Societe : Alten

set version=A9
echo Version : %version% 13.02.2013

set updateMacroFileName=C:\macros_alstom\UpdateMacroTCMS.exe

if exist %updateMacroFileName% goto CallUpdate
exit

:CallUpdate
echo call %updateMacroFileName% onStartup %version%
::pause
call %updateMacroFileName% onStartup %version%