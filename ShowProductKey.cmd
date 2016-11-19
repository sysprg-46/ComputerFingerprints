@echo off
set "PS=c:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
%PS% -windowstyle normal -NoExit -ExecutionPolicy BYPASS -File .\ShowProductKey.ps1
