@echo off
cd /d %~dp0

cscript vbac.wsf decombine /binary ..\bin /source ..\src
