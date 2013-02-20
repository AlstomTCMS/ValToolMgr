@echo off

rem ************************************************************
rem Appelle le .bat sur réseau pour la mise à jour
rem ************************************************************
rem UpdateMacroTCMS.exe
rem Auteur du fichier: DLE
rem Societe : Alten
rem Version : A10 20.02.2013

set macroPath=C:\macros_alstom\
set settingsFileName=%macroPath%Application_Settings_File.MIESET
set updateFileFullPath=\\dom2.ad.sys\dfsbor1root\bor1_flo\DEP_Etudes\Tsysteme\Affaires\PRIMA EL2\Ctrl-cmd\Banc de Test\13_Macros\
set updateFileName=install_auto_macro_alstom_tcms_prima.exe
set startupUpdateFileName=UpdateMacroOnStartup.exe
::echo %settingsFileName% %updateFileFullPath%%updateFileName%

:: Vérifie si on est en mode manuel
echo argument de la fonction: %1 %2
::pause
IF %1 == manuel goto CallUpdate
::pause
IF %1 == onStartup goto ExistSettings

IF %1 == checkStartup goto ExistSettings
exit

:ExistSettings
IF exist %settingsFileName% goto IsAutoUpdate
echo Le fichier de configuration %settingsFileName% n'existe pas !
::pause
exit

:IsAutoUpdate
echo Le fichier de configuration %settingsFileName% existe
::pause

for /F "tokens=1,2 delims=|" %%a in ('findstr /I "AutoUpdate" %settingsFileName%') do set isAutoUpdate="%%b

::pause
echo Mise a jour automatique : %isAutoUpdate%

IF %1 == checkStartup goto CheckStartup

::::pause
if %isAutoUpdate% == "True" goto CallUpdate
::echo Pas de mise a jour automatique
::pause
exit

:CallUpdate
echo Appel du fichier "%updateFileFullPath%%updateFileName%" avec parametres:%1 %2
::::pause
call "%updateFileFullPath%%updateFileName%" %1 %2
exit


::-------------------------------------------------------------------------------------------------------------------------------------


:CheckStartup
@echo off
cls
:: met l'encodage qui permet aux chemins avec accents de passer (D:\Documents and Settings\e_dleona\Menu Démarrer\Programmes\Démarrage)
chcp 1250


:GetWindowsVersion
REM Check Windows Version
ver | findstr /i "5\.0\." > nul
IF %ERRORLEVEL% EQU 0 goto GetStartupPath_XP 
::ver_2000
ver | findstr /i "5\.1\." > nul
IF %ERRORLEVEL% EQU 0 goto GetStartupPath_XP
::ver_XP
ver | findstr /i "5\.2\." > nul
IF %ERRORLEVEL% EQU 0 goto GetStartupPath_XP
::ver_2003
ver | findstr /i "6\.0\." > nul
IF %ERRORLEVEL% EQU 0 goto GetStartupPath_XP
::ver_Vista
ver | findstr /i "6\.1\." > nul
IF %ERRORLEVEL% EQU 0 goto GetStartupPath_win7
goto warn_and_exit


::skip4 pour XP
::skip2 pour Win7
:GetStartupPath_XP
echo Recherche du dossier de demarrage sous XP
for /F "skip=4 tokens=2*" %%j in ('reg query "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v "Startup"') do set startupPath=%%k
echo dossier de demarrage: %startupPath%
goto CopyOrDeleteFiles

:GetStartupPath_win7
echo Recherche du dossier de demarrage sous Win7
for /F "skip=2 tokens=2*" %%j in ('reg query "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v "Startup"') do set startupPath=%%k
echo dossier de demarrage: %startupPath%
goto CopyOrDeleteFiles

:CopyOrDeleteFiles
::pause
:: Si on passe en mode autoupdate, on ajoute le fichier .bat dans le startup
if %isAutoUpdate% == "True" xcopy "%macroPath%%startupUpdateFileName%" "%startupPath%" /Y 

:: Si on enleve le mode autoupdate, on supprime le fichier .bat dans le startup
if %isAutoUpdate% == "False" echo y|del "%startupPath%\%startupUpdateFileName%">nul 
::pause