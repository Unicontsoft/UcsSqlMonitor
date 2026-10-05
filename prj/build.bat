@echo off
setlocal
set VbCodeLines="C:\work\BuildTools\VBCodeLines\VbCodeLines.exe"
for %%i in ("%~dp0.") do set prj_dir=%%~dpnxi
for %%i in ("%prj_dir%\..\src\.") do set src_dir=%%~dpnxi
for %%i in ("%prj_dir%\..\res\.") do set res_dir=%%~dpnxi
for %%i in ("%prj_dir%\..\bin\.") do set bin_dir=%%~dpnxi
for %%i in ("%prj_dir%\..\compile\.") do set compile_dir=%%~dpnxi
set Vb6="%ProgramFiles%\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist %Vb6% set Vb6="%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"

echo Copy sources to %compile_dir%...
if not exist "%compile_dir%" mkdir "%compile_dir%"
del /q "%compile_dir%\*.*" > nul 2>&1
for %%i in ("%src_dir%\*.bas" "%src_dir%\*.cls" "%src_dir%\*.ctl" "%src_dir%\*.ctx" "%src_dir%\*.frm" "%src_dir%\*.frx") do (copy "%%i" "%compile_dir%" > nul)
copy "%res_dir%\UcsSqlMonitor.res" "%compile_dir%" > nul
rem the project points to ..\src and ..\res while the copies sit next to it
powershell -NoProfile -Command "(Get-Content '%prj_dir%\UcsSqlMonitor.vbp') -replace '\.\.\\(src|res)\\', '' | Set-Content -Encoding Default '%compile_dir%\UcsSqlMonitor.vbp'"
echo Put lines to sources in %compile_dir%...
%VbCodeLines% "%compile_dir%\UcsSqlMonitor.vbp"
echo Compiling to %bin_dir%...
attrib -r "%bin_dir%\*.*"
%Vb6% /m "%compile_dir%\UcsSqlMonitor.vbp"
rd /s /q "%compile_dir%"
echo Done.
