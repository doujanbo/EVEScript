%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit
@echo off 
set "sd=D:\mouse"

cd/d "%sd%"
for /r %%a in (*.exe) do (
    netsh advfirewall firewall del rule name="阻止%%~nxa联网">nul 2>nul
    netsh advfirewall firewall add rule name="阻止%%~nxa联网" program=%%a action=block dir=out>nul
    echo;阻止%%~nxa联网
)
pause