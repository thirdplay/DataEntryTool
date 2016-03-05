@echo off
cd %~dp0

@rem テンプレートファイルをbinディレクトリにコピーする
if not exist bin mkdir bin
cd src
for /d %%A in (*.xlsm) do copy %%A\template.xlsm ..\bin\%%A >nul
cd ..

@rem 結合処理実行
cscript //nologo vbac.wsf combine
if not %ERRORLEVEL%==0 pause
