<#
    Rebuild res\UcsSqlMonitor.res from res\UcsSqlMonitor.rc.

        powershell -File res\make-res.ps1

    The .res carries the manifest (Common Controls 6.0, DPI aware) and the
    Fugue PNGs for the server tree, which GDI+ decodes at run time.

    Only needed after editing UcsSqlMonitor.manifest or the icons; the .res is
    committed.
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$rc = 'C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x86\rc.exe'
if (-not (Test-Path -LiteralPath $rc)) {
    $rc = 'C:\Program Files (x86)\Microsoft Visual Studio\Common\MSDev98\Bin\RC.EXE'
}
if (-not (Test-Path -LiteralPath $rc)) { Write-Error 'make-res: rc.exe not found' }

Push-Location $PSScriptRoot
try {
    & $rc /nologo /fo UcsSqlMonitor.res UcsSqlMonitor.rc
    if ($LASTEXITCODE) { Write-Error "make-res: rc.exe returned $LASTEXITCODE" }
} finally {
    Pop-Location
}
'rebuilt res\UcsSqlMonitor.res'
