@echo off
set BIN_DIR=bin
set BAK_DIR=bak%DATE:/=%
cd %~dp0

@rem 一時ディレクトリ作成
mkdir %BAK_DIR%
copy %BIN_DIR%\* %BAK_DIR% >nul

@rem 分離処理実行
cscript //nologo vbac.wsf decombine /binary %BAK_DIR%
cscript //nologo vbac.wsf clear /binary %BAK_DIR%

@rem bak/*.xlsmをsrcにコピー
cd %BAK_DIR%
for %%A in (*.xlsm) do copy %%A ..\src\%%A\template.xlsm >nul
cd ..

@rem 一時ディレクトリ削除
rmdir /s /q %BAK_DIR%
