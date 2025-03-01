@echo off

createlink.exe >NUL
if errorlevel 0 echo "createlink.exe bereit"

MD "%USERPROFILE%\Desktop\Test" >NUL

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

pause
