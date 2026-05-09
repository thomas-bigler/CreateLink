:: set encoding to UTF8
@chcp 65001 1>nul 2>nul
@echo off

createlink.exe 1>NUL 2>NUL
if errorlevel 1 echo "CreateLink.exe not found" && GOTO DONE

echo "CreateLink.exe is here"
  
:: ECHO.
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

:DONE
pause
