@echo off
cd %~dp0

@rem �e���v���[�g�t�@�C����bin�f�B���N�g���ɃR�s�[����
cd src
for /d %%A in (*.xlsm) do copy %%A\template.xlsm ..\bin\%%A >nul
cd ..

@rem �����������s
cscript //nologo vbac.wsf combine
