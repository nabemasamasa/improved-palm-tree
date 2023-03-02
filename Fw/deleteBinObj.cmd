@echo off
set SRC_EXIST="0"
for /D /R %%d in (*) do if "bin"=="%%~nxd" echo %%d && set SRC_EXIST="1"
for /D /R %%d in (*) do if "obj"=="%%~nxd" echo %%d && set SRC_EXIST="1"
if SRC_EXIST=="0" goto END

set /P INP="以上のフォルダを削除します. よろしいですか？ Y/N >"
if "%INP%"=="Y" goto CONTINUE
if "%INP%"=="y" goto CONTINUE
echo 中断します.
goto END

:CONTINUE
for /D /R %%d in (*) do if "bin"=="%%~nxd" rd /S /Q "%%d" && echo %%d を削除...
for /D /R %%d in (*) do if "obj"=="%%~nxd" rd /S /Q "%%d" && echo %%d を削除...

:END
exit /b
