@echo off
windres -o version.res version.rc
fpc.exe -Px86_64 CreateLink.pas