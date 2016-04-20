@echo off
cd %~dp0

@rem •ª—£ˆ—Às
cscript //nologo vbac.wsf decombine /vbaproj /template
if not %ERRORLEVEL%==0 pause
