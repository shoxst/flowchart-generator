@echo off
cd /d %~dp0

set /p input=xlsm�t�@�C����src�ȉ��̃R�[�h�ŏ㏑������܂��B^
��낵���ł����H(yes/no):

if not "%input%" == "yes" (
    echo �����𒆒f���܂���
    exit /b
)

cscript vbac.wsf combine /binary ..\bin /source ..\src
