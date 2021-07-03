@echo off
cd /d %~dp0

set /p input=xlsmファイルがsrc以下のコードで上書きされます。^
よろしいですか？(yes/no):

if not "%input%" == "yes" (
    echo 処理を中断しました
    exit /b
)

cscript vbac.wsf combine /binary ..\bin /source ..\src
