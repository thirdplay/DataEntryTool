@echo off
set BIN_DIR=bin
set BAK_DIR=bak%DATE:/=%
cd %~dp0

@rem �ꎞ�f�B���N�g���쐬
mkdir %BAK_DIR%
copy %BIN_DIR%\* %BAK_DIR% >nul

@rem �����������s
cscript //nologo vbac.wsf decombine /binary %BAK_DIR%
cscript //nologo vbac.wsf clear /binary %BAK_DIR%

@rem bak/*.xlsm��src�ɃR�s�[
cd %BAK_DIR%
for %%A in (*.xlsm) do copy %%A ..\src\%%A\template.xlsm >nul
cd ..

@rem �ꎞ�f�B���N�g���폜
rmdir /s /q %BAK_DIR%
