@echo off
cd %~dp0

@rem 分離処理実行
cscript //nologo vbac.wsf decombine /vbaproj /template
if not %ERRORLEVEL%==0 pause
