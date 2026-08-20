<#
.SYNOPSIS
    Diagnoses why the Abbreviator add-in does not load in Word.

.DESCRIPTION
    Checks the registry, then reproduces outside Word exactly what Word does
    when it loads a COM add-in: create the class, then ask the CCW for
    IUnknown, IDispatch, IDTExtensibility2 and IRibbonExtensibility.
    Any failure is reported with the real exception, which Word never shows.
    Finally prints the add-in log and related Windows event log entries.

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

function Show-Exception($ex, $indent = '    ') {
    while ($ex) {
        Write-Host ($indent + $ex.GetType().FullName + ': ' + $ex.Message) -ForegroundColor Red
        if ($ex -is [System.Runtime.InteropServices.COMException]) {
            Write-Host ($indent + 'HRESULT: 0x' + ('{0:X8}' -f $ex.HResult)) -ForegroundColor Red
        }
        $ex = $ex.InnerException
        $indent = $indent + '    '
    }
}

# ------------------------------------------------------------- registry ---

Section 'Word add-in key'
$addin = "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
if (Test-Path $addin) {
    $lb = (Get-ItemProperty $addin).LoadBehavior
    Write-Host "LoadBehavior = $lb"
    switch ($lb) {
        3       { Write-Host 'Loads at startup (expected).' -ForegroundColor Green }
        2       { Write-Host 'Word disabled the add-in after a load failure.' -ForegroundColor Red }
        default { Write-Host 'Not the expected value 3.' -ForegroundColor Yellow }
    }
} else {
    Write-Host 'NOT REGISTERED - run tools\install.ps1' -ForegroundColor Red
}

Section 'COM registration'
$dllPath = $null
foreach ($classes in @('HKCU:\Software\Classes', 'HKCU:\Software\Classes\Wow6432Node')) {
    $inproc = "$classes\CLSID\$Clsid\InprocServer32"
    Write-Host $inproc
    if (Test-Path $inproc) {
        $p = Get-ItemProperty $inproc
        Write-Host ('    server   : ' + $p.'(default)')
        Write-Host ('    class    : ' + $p.Class)
        Write-Host ('    assembly : ' + $p.Assembly)
        Write-Host ('    runtime  : ' + $p.RuntimeVersion)
        Write-Host ('    codebase : ' + $p.CodeBase)

        $file = $p.CodeBase -replace '^file:///', '' -replace '/', '\'
        if (Test-Path $file) {
            $info = Get-Item $file
            Write-Host ('    file     : OK, ' + $info.Length + ' bytes, ' + $info.LastWriteTime) -ForegroundColor Green
            if (-not $dllPath) { $dllPath = $file }
        } else {
            Write-Host '    file     : MISSING - rebuild or re-run install.ps1' -ForegroundColor Red
        }
    } else {
        Write-Host '    (absent)' -ForegroundColor Yellow
    }
}

# --------------------------------------------------------- CCW self-test ---

Section 'COM self-test (what Word does)'

if (-not $dllPath) {
    $guess = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\Abbreviator\bin\Release\Abbreviator.dll'
    if (Test-Path $guess) { $dllPath = $guess }
}

if (-not $dllPath) {
    Write-Host 'Abbreviator.dll not found - build the project first.' -ForegroundColor Red
} else {
    Write-Host "Assembly: $dllPath"
    Write-Host ("PowerShell process: " + $(if ([IntPtr]::Size -eq 8) { '64-bit' } else { '32-bit' }))
    Write-Host ''

    $obj = $null
    try {
        $asm = [System.Reflection.Assembly]::LoadFrom($dllPath)
        Write-Host '1. Load assembly            : OK' -ForegroundColor Green

        $type = $asm.GetType('Abbreviator.Connect', $true)
        $obj  = [System.Activator]::CreateInstance($type)
        Write-Host '2. Create managed instance  : OK' -ForegroundColor Green
    }
    catch {
        Write-Host '   FAILED' -ForegroundColor Red
        Show-Exception $_.Exception
    }

    if ($obj) {
        # IUnknown: builds the COM callable wrapper.
        try {
            $unk = [System.Runtime.InteropServices.Marshal]::GetIUnknownForObject($obj)
            Write-Host '3. CCW / IUnknown           : OK' -ForegroundColor Green
            [void][System.Runtime.InteropServices.Marshal]::Release($unk)
        } catch {
            Write-Host '3. CCW / IUnknown           : FAILED' -ForegroundColor Red
            Show-Exception $_.Exception
        }

        # IDispatch: this is what generates the class interface. Ribbon and
        # context menu callbacks are dispatched through it by name.
        try {
            $disp = [System.Runtime.InteropServices.Marshal]::GetIDispatchForObject($obj)
            Write-Host '4. IDispatch (class iface)  : OK' -ForegroundColor Green
            [void][System.Runtime.InteropServices.Marshal]::Release($disp)
        } catch {
            Write-Host '4. IDispatch (class iface)  : FAILED' -ForegroundColor Red
            Write-Host '   Ribbon callbacks cannot work without this.' -ForegroundColor Yellow
            Show-Exception $_.Exception
        }

        # The interfaces Word queries for. A failure here is exactly the case
        # where the log has "object created" and nothing else.
        $step = 5
        foreach ($name in @('Abbreviator.Interop.IDTExtensibility2',
                            'Abbreviator.Interop.IRibbonExtensibility')) {
            $short = $name.Split('.')[-1]
            try {
                $it  = $asm.GetType($name, $true)
                $ptr = [System.Runtime.InteropServices.Marshal]::GetComInterfaceForObject($obj, $it)
                Write-Host ("$step. QueryInterface " + $short.PadRight(11) + ': OK') -ForegroundColor Green
                [void][System.Runtime.InteropServices.Marshal]::Release($ptr)
            } catch {
                Write-Host ("$step. QueryInterface " + $short.PadRight(11) + ': FAILED') -ForegroundColor Red
                Write-Host '   This is why Word refuses to load the add-in.' -ForegroundColor Yellow
                Show-Exception $_.Exception
            }
            $step++
        }
    }

    # CoCreateInstance through the registration, the way Word starts.
    try {
        $com = New-Object -ComObject $ProgId
        Write-Host '7. CoCreateInstance by ProgId: OK' -ForegroundColor Green
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($com)
    } catch {
        Write-Host '7. CoCreateInstance by ProgId: FAILED' -ForegroundColor Red
        Show-Exception $_.Exception
    }
}

# ---------------------------------------------------------- disabled items ---

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

# ------------------------------------------------------------------- logs ---

Section 'Add-in log'
if (Test-Path $Log) {
    Write-Host $Log
    Write-Host ''
    Get-Content $Log -Tail $LogLines
} else {
    Write-Host "$Log does not exist." -ForegroundColor Red
    Write-Host 'No add-in code ever ran - look at the COM registration above.' -ForegroundColor Yellow
}

Section 'Windows event log (last 24h)'
$since = (Get-Date).AddDays(-1)
$events = Get-EventLog -LogName Application -After $since |
          Where-Object { $_.Source -match 'Office|\.NET Runtime|Application Error|Application Hang' } |
          Select-Object -First 15
if ($events) {
    foreach ($e in $events) {
        Write-Host ''
        Write-Host ("[$($e.TimeGenerated)] $($e.EntryType) - $($e.Source)") -ForegroundColor Yellow
        $text = $e.Message -split "`n" | Select-Object -First 6
        $text | ForEach-Object { Write-Host ('    ' + $_.TrimEnd()) }
    }
} else {
    Write-Host 'No related entries.' -ForegroundColor Green
}

Write-Host ''
