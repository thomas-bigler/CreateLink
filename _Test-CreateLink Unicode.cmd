@echo off
:: set encoding to UTF8 
chcp 65001
echo off

ECHO.
echo Checking CreateLink.exe . . .
createlink.exe >NUL
if errorlevel 0 echo "createlink.exe bereit"

ECHO.
MD "%USERPROFILE%\Desktop\Test" 1>nul 2>nul

CreateLink "%USERPROFILE%\Desktop\Test\IrfanView 64 Bit - Юнікод!.lnk" "c:\Program Files\IrfanView\i_view64.exe"

CreateLink "%USERPROFILE%\Desktop\Test\Far Manager (Standard) - Юнікод!.lnk" "c:\Program Files\Far Manager\Far.exe"

CreateLink "%USERPROFILE%\Desktop\Test\Far Manager (Maximized) - Юнікод!.lnk"^
 "c:\Program Files\Far Manager\Far.exe"^
 ""^
 "c:\Program Files\Far Manager"^
 "Classical File Manager"^
 "c:\Program Files\Far Manager\Far.exe"^
 2^
 3


pause
