@echo off
cd %~dp0

@rem �����������s
cscript //nologo vbac.wsf decombine /template
if not %ERRORLEVEL%==0 pause
