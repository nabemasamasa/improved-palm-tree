@echo off
set SRC_EXIST="0"
for /D /R %%d in (*) do if "bin"=="%%~nxd" echo %%d && set SRC_EXIST="1"
for /D /R %%d in (*) do if "obj"=="%%~nxd" echo %%d && set SRC_EXIST="1"
if SRC_EXIST=="0" goto END

set /P INP="�ȏ�̃t�H���_���폜���܂�. ��낵���ł����H Y/N >"
if "%INP%"=="Y" goto CONTINUE
if "%INP%"=="y" goto CONTINUE
echo ���f���܂�.
goto END

:CONTINUE
for /D /R %%d in (*) do if "bin"=="%%~nxd" rd /S /Q "%%d" && echo %%d ���폜...
for /D /R %%d in (*) do if "obj"=="%%~nxd" rd /S /Q "%%d" && echo %%d ���폜...

:END
exit /b
