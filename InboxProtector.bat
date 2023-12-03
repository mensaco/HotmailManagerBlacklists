@REM @echo off

@REM InboxProtector.bat

@echo off
setlocal

REM Get the directory where the batch script is located
set "ScriptDir=%~dp0"

REM Construct the path to the PowerShell script
set "PowerShellScript=%ScriptDir%InboxProtector.ps1"

REM Execute the PowerShell script
powershell -ExecutionPolicy Bypass -File "%PowerShellScript%"

endlocal