<#
.SYNOPSIS
    Собирает и регистрирует надстройку Abbreviator для текущего пользователя.

.DESCRIPTION
    Регистрация выполняется только в HKCU, права администратора не нужны.
    Записываются два набора ключей — обычный и Wow6432Node, чтобы надстройка
    подхватилась и 64-разрядным, и 32-разрядным Word.

.PARAMETER Dll
    Путь к уже собранной Abbreviator.dll. Если не задан, проект собирается
    командой dotnet build -c Release.

.PARAMETER SkipBuild
    Не собирать проект, использовать то, что уже лежит в bin\Release.

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
$Description  = 'Проверка аббревиатур и перечня принятых сокращений'

$root    = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root 'src\Abbreviator\Abbreviator.csproj'
$output  = Join-Path $root 'src\Abbreviator\bin\Release\Abbreviator.dll'

# ---------------------------------------------------------------- сборка ---

if (-not $Dll) {
    if (-not $SkipBuild) {
        Write-Host 'Сборка проекта...' -ForegroundColor Cyan
        & dotnet build $project -c Release
        if ($LASTEXITCODE -ne 0) { throw "Сборка завершилась с ошибкой ($LASTEXITCODE)." }
    }
    $Dll = $output
}

$Dll = (Resolve-Path $Dll).Path
if (-not (Test-Path $Dll)) { throw "Не найден файл $Dll" }

Write-Host "Регистрируется: $Dll" -ForegroundColor Cyan

# ------------------------------------------------------------ регистрация ---

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

# Обе ветки: 64-разрядный и 32-разрядный Word.
$classRoots = @(
    'HKCU:\Software\Classes',
    'HKCU:\Software\Classes\Wow6432Node'
)

foreach ($classes in $classRoots) {
    Set-Key "$classes\$ProgId"                      @{ '(default)' = $FriendlyName }
    Set-Key "$classes\$ProgId\CLSID"                @{ '(default)' = $Clsid }
    Set-Key "$classes\CLSID\$Clsid"                 @{ '(default)' = $ClassName }
    Set-Key "$classes\CLSID\$Clsid\ProgId"          @{ '(default)' = $ProgId }
    Set-Key "$classes\CLSID\$Clsid\InprocServer32"  $inproc
    Set-Key "$classes\CLSID\$Clsid\InprocServer32\1.0.0.0" $inproc
}

# Ключ надстройки Word. LoadBehavior=3 — загружать при старте.
$addin = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
Set-Key $addin @{
    'FriendlyName'    = $FriendlyName
    'Description'     = $Description
    'LoadBehavior'    = 3
    'CommandLineSafe' = 0
}

Write-Host ''
Write-Host 'Готово. Запустите Word — на ленте появится вкладка "Abbreviator".' -ForegroundColor Green
Write-Host 'Если вкладки нет: Файл → Параметры → Надстройки → Надстройки COM → Перейти,' -ForegroundColor Yellow
Write-Host 'и поставьте галочку у Abbreviator.' -ForegroundColor Yellow
