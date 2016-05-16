@echo off
color 1F
echo.
::the bat file should be same path with *.ps1 file
#-noexit
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -noexit  -windowstyle hidden -NoProfile -ExecutionPolicy Bypass -Command  "& '%~dp0\Monitor.ps1'"
:EOF
echo Waiting seconds
timeout /t 10 /nobreak > NUL