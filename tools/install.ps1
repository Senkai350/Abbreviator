<#
.SYNOPSIS
    Builds and registers the Abbreviator Word add-in for the current user.

.DESCRIPTION
    Registration is written to HKCU only, so administrator rights are not
    required. Two sets of keys are written - the normal one and Wow6432Node -
    so that both 64-bit and 32-bit Word pick the add-in up.

.PARAMETER Dll
    Path to an already built Abbreviator.dll. When omitted, the project is
    built with: dotnet build -c Release

.PARAMETER SkipBuild
    Do not build; use whatever is already in bin\Release.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\install.ps1
#>

[CmdletBinding()]
param(
    [string] $Dll,
    [switch] $SkipBuild
)

$ErrorActionPreference = 'Stop'

$Clsid        = '{FE63C2BE-14F3-45D9-BE38-98A925BDF989}'
$ProgId       = 'Abbreviator.Connect'
$AssemblyName = 'Abbreviator, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$ClassName    = 'Abbreviator.Connect'
$Runtime      = 'v4.0.30319'
$FriendlyName = 'Abbreviator'
$Description  = 'Abbreviation checker for Word'

$root    = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root 'src\Abbreviator\Abbreviator.csproj'
$output  = Join-Path $root 'src\Abbreviator\bin\Release\Abbreviator.dll'

# ------------------------------------------------------------------- build ---

if (-not $Dll) {
    if (-not $SkipBuild) {
        Write-Host 'Building...' -ForegroundColor Cyan
        & dotnet build $project -c Release
        if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE." }
    }
    $Dll = $output
}

if (-not (Test-Path $Dll)) { throw "File not found: $Dll" }
$Dll = (Resolve-Path $Dll).Path

Write-Host "Registering: $Dll" -ForegroundColor Cyan

# -------------------------------------------------------------- registration ---

function Set-Key {
    param([string] $Path, [hashtable] $Values)

    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    foreach ($name in $Values.Keys) {
        $value = $Values[$name]
        $type  = if ($value -is [int]) { 'DWord' } else { 'String' }
        New-ItemProperty -Path $Path -Name $name -Value $value -PropertyType $type -Force | Out-Null
    }
}

$codeBase = 'file:///' + ($Dll -replace '\\', '/')

$inproc = @{
    '(default)'      = 'mscoree.dll'
    'ThreadingModel' = 'Both'
    'Class'          = $ClassName
    'Assembly'       = $AssemblyName
    'RuntimeVersion' = $Runtime
    'CodeBase'       = $codeBase
}

# Both views, so 64-bit and 32-bit Word can load the add-in.
$classRoots = @(
    'HKCU:\Software\Classes',
    'HKCU:\Software\Classes\Wow6432Node'
)

foreach ($classes in $classRoots) {
    Set-Key "$classes\$ProgId"                             @{ '(default)' = $FriendlyName }
    Set-Key "$classes\$ProgId\CLSID"                       @{ '(default)' = $Clsid }
    Set-Key "$classes\CLSID\$Clsid"                        @{ '(default)' = $ClassName }
    Set-Key "$classes\CLSID\$Clsid\ProgId"                 @{ '(default)' = $ProgId }
    Set-Key "$classes\CLSID\$Clsid\InprocServer32"         $inproc
    Set-Key "$classes\CLSID\$Clsid\InprocServer32\1.0.0.0" $inproc
}

# Word add-in key. LoadBehavior = 3 means "load at startup".
$addin = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
Set-Key $addin @{
    'FriendlyName'    = $FriendlyName
    'Description'     = $Description
    'LoadBehavior'    = 3
    'CommandLineSafe' = 0
}

Write-Host ''
Write-Host 'Done. Start Word - the ribbon tab "Abbreviator" should appear.' -ForegroundColor Green
Write-Host 'If it does not: File > Options > Add-ins > Manage: COM Add-ins > Go,' -ForegroundColor Yellow
Write-Host 'then tick the Abbreviator checkbox.' -ForegroundColor Yellow
