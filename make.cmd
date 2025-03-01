@echo off
windres -o version.res version.rc
fpc.exe CreateLink.pas