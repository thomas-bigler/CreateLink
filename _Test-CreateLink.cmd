:: set encoding to UTF8
@chcp 65001 1>nul 2>nul
@echo off

createlink.exe 1>NUL 2>NUL
if errorlevel 1 echo "CreateLink.exe not found" && GOTO DONE
echo "CreateLink.exe is here"

MD "%USERPROFILE%\Desktop\Test" 2>NUL

CreateLink "%USERPROFILE%\Desktop\Test\Far Manager (Standard).lnk" "c:\Program Files\Far Manager\Far.exe"

CreateLink "%USERPROFILE%\Desktop\Test\Far Manager (Maximized).lnk"^
 "c:\Program Files\Far Manager\Far.exe"^
 ""^
 "c:\Program Files\Far Manager"^
 "Classical File Manager"^
 "c:\Program Files\Far Manager\Far.exe"^
 2^
 3

CreateLink "%USERPROFILE%\Desktop\Test\IrfanView 64 Bit.lnk" "c:\Program Files\IrfanView\i_view64.exe"

CreateLink "%USERPROFILE%\Desktop\Test\Test BAD.lnk" "NUL"
:DONE
pause
