@echo off
setlocal
cd /d %~dp0
set MKTYPLIB=mktyplib.exe
where %MKTYPLIB% >nul 2>&1 || set MKTYPLIB="C:\Program Files (x86)\Microsoft Visual Studio\VC98\Bin\MKTYPLIB.EXE"
%MKTYPLIB% /nologo /win32 /nocpp UcsSqlMonitor.odl /tlb UcsSqlMonitor.tlb
