<#
.SYNOPSIS
    Prints everything needed to diagnose why the Abbreviator add-in does not load.

.DESCRIPTION
    Shows the registry state (CLSID, ProgId, LoadBehavior), whether the DLL is
    where the registration points, whether Word has the add-in in its disabled
    list, and the tail of the add-in log.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\diagnose.ps1
#>

[CmdletBinding()]
param([int] $LogLines = 60)

$ErrorActionPreference = 'SilentlyContinue'

$Clsid  = '{FE63C2BE-14F3-45D9-BE38-98A925BDF989}'
$ProgId = 'Abbreviator.Connect'
$Log    = Join-Path $env:APPDATA 'Abbreviator\abbreviator.log'

function Section($title) {
    Write-Host ''
    Write-Host "--- $title " -ForegroundColor Cyan -NoNewline
    Write-Host ('-' * [Math]::Max(0, 60 - $title.Length)) -ForegroundColor Cyan
}

Section 'Word add-in key'
$addin = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
if (Test-Path $addin) {
    Get-ItemProperty $addin | Format-List FriendlyName, Description, LoadBehavior
    $lb = (Get-ItemProperty $addin).LoadBehavior
    switch ($lb) {
        3       { Write-Host 'LoadBehavior 3 - loads at startup (expected).' -ForegroundColor Green }
        2       { Write-Host 'LoadBehavior 2 - Word disabled the add-in after a load failure.' -ForegroundColor Red }
        default { Write-Host "LoadBehavior $lb - not the expected value 3." -ForegroundColor Yellow }
    }
} else {
    Write-Host 'NOT REGISTERED - run tools\install.ps1' -ForegroundColor Red
}

Section 'COM registration'
foreach ($classes in @('HKCU:\Software\Classes', 'HKCU:\Software\Classes\Wow6432Node')) {
    $inproc = "$classes\CLSID\$Clsid\InprocServer32"
    Write-Host "$inproc"
    if (Test-Path $inproc) {
        $p = Get-ItemProperty $inproc
        Write-Host ("    server   : " + $p.'(default)')
        Write-Host ("    class    : " + $p.Class)
        Write-Host ("    assembly : " + $p.Assembly)
        Write-Host ("    runtime  : " + $p.RuntimeVersion)
        Write-Host ("    codebase : " + $p.CodeBase)

        $file = $p.CodeBase -replace '^file:///', '' -replace '/', '\'
        if (Test-Path $file) {
            $info = Get-Item $file
            Write-Host ("    file     : OK, " + $info.Length + " bytes, " + $info.LastWriteTime) -ForegroundColor Green
        } else {
            Write-Host '    file     : MISSING - rebuild or re-run install.ps1' -ForegroundColor Red
        }
    } else {
        Write-Host '    (absent)' -ForegroundColor Yellow
    }
}

Section 'Word disabled items'
$found = $false
Get-ChildItem 'HKCU:\Software\Microsoft\Office' |
    Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
    ForEach-Object {
        foreach ($leaf in @('DisabledItems', 'CrashingAddinList')) {
            $path = Join-Path $_.PSPath "Word\Resiliency\$leaf"
            if (Test-Path $path) {
                Write-Host "$($_.PSChildName) Word\Resiliency\$leaf exists" -ForegroundColor Red
                $found = $true
            }
        }
    }
if (-not $found) { Write-Host 'Nothing disabled.' -ForegroundColor Green }
else { Write-Host 'Re-run install.ps1 to clear these, then restart Word.' -ForegroundColor Yellow }

Section 'Add-in log'
if (Test-Path $Log) {
    Write-Host "$Log"
    Write-Host ''
    Get-Content $Log -Tail $LogLines
} else {
    Write-Host "$Log does not exist." -ForegroundColor Red
    Write-Host 'The add-in never got as far as running any code - the problem is' -ForegroundColor Yellow
    Write-Host 'in the COM registration or the assembly could not be loaded.' -ForegroundColor Yellow
}

Write-Host ''
