@echo off
cls
:: met l'encodage qui permet aux chemins avec accents de passer (D:\Documents and Settings\e_dleona\Menu D�marrer\Programmes\D�marrage)
chcp 1250

rem ****************************************
rem Installation automatique de la macro
rem ****************************************
rem install_auto_macro_alstom_tcms_prima.exe
rem Auteur du fichier: DLE
rem Societe : Alten
rem version : A7 08.02.2013
set versionServeur=A7


:checkIsToUpdate
echo. Version installee: %2		Version serveur: %versionServeur%
::pause
if "%2" == "" goto initPath
if "%2" == "version" goto update
::pause
if %2 gtr %versionServeur% exit 
::echo Version %2 plus grand que %versionServeur%

rem On ne met � jour que si la version install�e est inf�rieure � celle sur serveur
if %2 Lss %versionServeur% goto update
::pause
if %2 Equ %versionServeur% goto alreadyUpdate
::Version %2 aussi grand que A5

:alreadyUpdate 
echo La macro est deja a jour ! 
if %1 == manuel pause
exit

:update 
echo La version installee n'est pas a jour.
echo Lancement de la mise a jour :
::pause

:initPath
set networkPath=I:\DEP_Etudes\Tsysteme\Affaires\PRIMA EL2\Ctrl-cmd\Banc de Test\13_Macros\
set localPath=C:\macros_alstom\

:deleteBat
rem Efface les anciens .bat
echo y|del "%localPath%\UpdateMacroTCMS.bat">nul
echo y|del "%localPath%\UpdateMacroOnStartup.bat">nul

:CopyFiles
echo.
rem Installation de la macro (copie le fichier sur le r�seau vers un dossier en local)
xcopy "%networkPath%Functions_PrimaELII_2-A0.xlam" %localPath% /Y
rem Copie du fichier des references (source de donn�es pour la macro)
xcopy "%networkPath%Ref_PrimaELII_2-A3.xls" %localPath% /Y 
rem Copie du .bat appel� pour la MAJ
xcopy "%networkPath%UpdateMacroTCMS.exe" %localPath% /Y 
rem Copie de sauvegarde du .bat pour une mise � jour auto en d�but de session utilisateur
xcopy "%networkPath%UpdateMacroOnStartup.exe" %localPath% /Y 

if NOT "%1" == "onStartup" goto TestExcelIsLaunched
goto Install

:TestExcelIsLaunched
set ExcelIsRunning=0
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
if "%ERRORLEVEL%"=="0" set ExcelIsRunning=1
::echo Excel is running
echo ExcelIsRunning %ExcelIsRunning%
::pause

:Install
rem Lancement de l'installation de la macro par le fichier Excel et attend la fermeture d'Excel 
::rem Ne pas oublier de mettre dans le fichier l'action permet de quitter excel
echo Lancement de l'installation de la macro par le fichier Excel "%networkPath%install_macro_excel.xlsm"
start /wait "Installation de la macro" "%networkPath%install_macro_excel.xlsm"

if NOT "%1" == "onStartup" goto IsExcelToRelaunch
exit

:IsExcelToRelaunch
echo IsExcelToRelaunch %ExcelIsRunning%
::pause
if %ExcelIsRunning% == 1 goto RelaunchExcel
exit

:RelaunchExcel
echo Relaunch Excel
Start excel.exe
::pause
exit


:warn_and_exit
echo Machine OS cannot be determined.
pause
exit