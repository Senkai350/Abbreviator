<#
.SYNOPSIS
    Removes the Abbreviator Word add-in registration for the current user.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'SilentlyContinue'

$Clsid  = '{FE63C2BE-14F3-45D9-BE38-98A925BDF989}'
$ProgId = 'Abbreviator.Connect'

$paths = @(
    "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId",
    "HKCU:\Software\Classes\$ProgId",
    "HKCU:\Software\Classes\CLSID\$Clsid",
    "HKCU:\Software\Classes\Wow6432Node\$ProgId",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$Clsid"
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force
        Write-Host "Removed: $path"
    }
}

Write-Host ''
Write-Host 'Add-in unregistered. Restart Word.' -ForegroundColor Green
Write-Host 'Settings are still in %APPDATA%\Abbreviator - delete the folder if not needed.' -ForegroundColor Yellow
