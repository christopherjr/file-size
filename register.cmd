@echo off

rem Simple re-registration helper.  Pass r or R to register the release version,
rem while unregistering the debug version.
rem Pass no parameters to register the debug version, while unregistering the
rem release version.

set REG_CONFIG=Debug
set UNREG_CONFIG=Release

if "%1"=="r" set REG_CONFIG=Release
if "%1"=="R" set REG_CONFIG=Release

if "%REG_CONFIG%"=="Release" set UNREG_CONFIG=Debug

rem Unregister and re-register the 32-bit extension
C:\Windows\SYSWOW64\regsvr32.exe /u .\%UNREG_CONFIG%\StatusBarExt.dll
C:\Windows\SYSWOW64\regsvr32.exe .\%REG_CONFIG%\StatusBarExt.dll
rem Unregister and re-register the 64-bit extension
C:\Windows\System32\regsvr32.exe /u .\x64\%UNREG_CONFIG%\StatusBarExt.dll
C:\Windows\System32\regsvr32.exe .\x64\%REG_CONFIG%\StatusBarExt.dll
