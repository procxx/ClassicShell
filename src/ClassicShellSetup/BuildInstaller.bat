@REM !!!!! CHANGE THE GUIDS WHEN CHANGING THE VERSION !!!!!
SET CS_VERSION=3.6.4
SET CS_VERSION_STR=3_6_4
SET CS_VERSION_NUM=30604
SET CS_GUID32=64AB944D-7A0A-44bd-8F99-E7B3007DF02A
SET CS_GUID64=01489E0F-02FB-4154-B927-5FED88353702

@SET CS_ERROR=0

REM ********* Build 32-bit solution
"%VS90COMNTOOLS%..\IDE\devenv.com" ..\ClassicShell.sln /rebuild "Setup|Win32"
@if ERRORLEVEL 1 goto end


REM ********* Build 64-bit solution
"%VS90COMNTOOLS%..\IDE\devenv.com" ..\ClassicShell.sln /rebuild "Setup|x64"
@if ERRORLEVEL 1 goto end


REM ********* Build Help
hhc ..\Docs\Help\ClassicShell.hhp
@REM looks like hhc returns 0 for error, >0 for success
@if NOT ERRORLEVEL 1 goto end


REM ********* Build Ini Checksums
start /wait SetupHelper\Release\SetupHelper.exe crc ..\ClassicExplorer ..\ClassicStartMenu
@if ERRORLEVEL 1 goto end

REM ********* Make en-US.dll
cd ..
start /wait ClassicShellSetup\SetupHelper\Release\SetupHelper.exe makeEN ClassicExplorer\Setup\ClassicExplorer32.dll ClassicStartMenu\Setup\ClassicStartMenuDLL.dll ClassicIE9\Setup\ClassicIE9DLL_32.dll ClassicShellUpdate\Release\ClassicShellUpdate.exe
@if ERRORLEVEL 1 goto end

start /wait ClassicShellSetup\LocalizeCS\Release\LocalizeCS.exe extract en-US.dll en-US.csv

cd ClassicShellSetup

md Temp
del /Q Temp\*.*

@if not exist ..\Localization\Russian\ClassicShellText-ru-RU.wxl goto english

REM **************************** Russian

REM ********* Build 32-bit MSI
candle ClassicShellSetup.wxs -out Temp\ClassicShellSetup32.wixobj -ext WixUIExtension -ext WixUtilExtension -dx64=0 -dlang=ru
@if ERRORLEVEL 1 goto end

@REM We need to suppress ICE38 and ICE43 because they apply only to per-user installation. We only support per-machine installs
light Temp\ClassicShellSetup32.wixobj -out Temp\ClassicShellSetup32.msi -ext WixUIExtension -ext WixUtilExtension -loc ..\Localization\Russian\ClassicShellText-ru-RU.wxl -loc ..\Localization\Russian\WixUI_ru-RU.wxl -sice:ICE38 -sice:ICE43
@if ERRORLEVEL 1 goto end


REM ********* Build 64-bit MSI
candle ClassicShellSetup.wxs -out Temp\ClassicShellSetup64.wixobj -ext WixUIExtension -ext WixUtilExtension -dx64=1 -dlang=ru
@if ERRORLEVEL 1 goto end

@REM We need to suppress ICE38 and ICE43 because they apply only to per-user installation. We only support per-machine installs
light Temp\ClassicShellSetup64.wixobj -out Temp\ClassicShellSetup64.msi -ext WixUIExtension -ext WixUtilExtension -loc ..\Localization\Russian\ClassicShellText-ru-RU.wxl -loc ..\Localization\Russian\WixUI_ru-RU.wxl -sice:ICE38 -sice:ICE43
@if ERRORLEVEL 1 goto end


REM ********* Build MSI Checksums
start /wait SetupHelper\Release\SetupHelper.exe crcmsi Temp
@if ERRORLEVEL 1 goto end

REM ********* Build bootstrapper
"%VS90COMNTOOLS%..\IDE\devenv.com" ClassicShellSetup.sln /rebuild "Release|Win32"
@if ERRORLEVEL 1 goto end
del Release\ClassicShellSetup-ru.exe
ren Release\ClassicShellSetup.exe ClassicShellSetup-ru.exe


:english

REM **************************** English

REM ********* Build 32-bit MSI
candle ClassicShellSetup.wxs -out Temp\ClassicShellSetup32.wixobj -ext WixUIExtension -ext WixUtilExtension -dx64=0 -dlang=en
@if ERRORLEVEL 1 goto end

@REM We need to suppress ICE38 and ICE43 because they apply only to per-user installation. We only support per-machine installs
light Temp\ClassicShellSetup32.wixobj -out Temp\ClassicShellSetup32.msi -ext WixUIExtension -ext WixUtilExtension -loc ClassicShellText-en-US.wxl -sice:ICE38 -sice:ICE43
@if ERRORLEVEL 1 goto end


REM ********* Build 64-bit MSI
candle ClassicShellSetup.wxs -out Temp\ClassicShellSetup64.wixobj -ext WixUIExtension -ext WixUtilExtension -dx64=1 -dlang=en
@if ERRORLEVEL 1 goto end

@REM We need to suppress ICE38 and ICE43 because they apply only to per-user installation. We only support per-machine installs
light Temp\ClassicShellSetup64.wixobj -out Temp\ClassicShellSetup64.msi -ext WixUIExtension -ext WixUtilExtension -loc ClassicShellText-en-US.wxl -sice:ICE38 -sice:ICE43
@if ERRORLEVEL 1 goto end


REM ********* Build MSI Checksums
start /wait SetupHelper\Release\SetupHelper.exe crcmsi Temp
@if ERRORLEVEL 1 goto end

REM ********* Build bootstrapper
"%VS90COMNTOOLS%..\IDE\devenv.com" ClassicShellSetup.sln /rebuild "Release|Win32"
@if ERRORLEVEL 1 goto end














@goto EOF
:end
@SET CS_ERROR=1
pause
