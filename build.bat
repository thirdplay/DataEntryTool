@echo off
cd %~dp0

@rem 結合処理実行
cscript //nologo vbac.wsf combine /vbaproj /template
if not %ERRORLEVEL%==0 pause
