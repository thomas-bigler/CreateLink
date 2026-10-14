@echo off
ECHO CreateLink examples
ECHO.

CreateLink "%USERPROFILE%\Desktop\Notepad.lnk" "%SYSTEMROOT%\Notepad.exe"

CreateLink "%USERPROFILE%\Desktop\On-Screen Keyboard.lnk"^
 "%SYSTEMROOT%\System32\osk.exe"^
 ""^
 "%SYSTEMROOT%\System32"^
 "On-Screen Keyboard can be used instead of a physical keyboard"^
 "%SYSTEMROOT%\System32\osk.exe"^
 ""^
 7

CreateLink "%USERPROFILE%\Desktop\Far Manager.lnk"^
 "%ProgramW6432%\Far Manager\Far.exe"^
 ""^
 "%ProgramW6432%\Far Manager"^
 "Classical File Manager"^
 "%ProgramW6432%\Far Manager\Far.exe"^
 2^
 3

PAUSE
