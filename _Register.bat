@echo off
cd %~dp0
%~d0

echo Please make sure to run As Administrator or the registration will not work
pause

echo:
PowerShell.exe -ExecutionPolicy ByPass -File RegisterOracle12.2.ps1
pause