@REM !!!!! CHANGE THE GUIDS WHEN CHANGING THE VERSION !!!!!
SET CS_VERSION=3.2.0
SET CS_VERSION_STR=3_2_0
SET CS_VERSION_NUM=30101
SET CS_GUID32=101A3855-EB37-4760-9DCB-BA04A28E0F75
SET CS_GUID64=1EEF5C7E-C371-431D-A507-8C5B46EED7B4

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
start /wait ClassicShellSetup\SetupHelper\Release\SetupHelper.exe makeEN ClassicExplorer\Setup\ClassicExplorer32.dll ClassicStartMenu\Setup\ClassicStartMenuDLL.dll ClassicIE9\Setup\ClassicIE9DLL_32.dll
@if ERRORLEVEL 1 goto end

start /wait ClassicShellSetup\LocalizeCS\Release\LocalizeCS.exe extract en-US.dll en-US.csv

cd ClassicShellSetup

md Temp
del /Q Temp\*.*

REM ********* Build 32-bit MSI
candle ClassicShellSetup.wxs -out Temp\ClassicShellSetup32.wixobj -ext WixUIExtension -ext WixUtilExtension -dx64=0
@if ERRORLEVEL 1 goto end

@REM We need to suppress ICE38 and ICE43 because they apply only to per-user installation. We only support per-machine installs
light Temp\ClassicShellSetup32.wixobj -out Temp\ClassicShellSetup32.msi -ext WixUIExtension -ext WixUtilExtension -loc ClassicShellText-en-US.wxl -sice:ICE38 -sice:ICE43
@if ERRORLEVEL 1 goto end


REM ********* Build 64-bit MSI
candle ClassicShellSetup.wxs -out Temp\ClassicShellSetup64.wixobj -ext WixUIExtension -ext WixUtilExtension -dx64=1
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
