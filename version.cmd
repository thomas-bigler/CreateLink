@echo off
ECHO.
ECHO  WindRes^: Compiling version.rc to version.res . . .
windres -o version.res version.rc
ECHO  Done.
ECHO.
ping -n 3 localhost >NUL