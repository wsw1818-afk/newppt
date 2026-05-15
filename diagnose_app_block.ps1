param(
    [Parameter(Mandatory = $true)]
    [string]$AppPath,

    [string]$RunStartTest = "0"
)

$ErrorActionPreference = "Continue"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$desktop = [Environment]::GetFolderPath("Desktop")
$report = Join-Path $desktop ("DocExtractor_Security_Diagnosis_" + $stamp + ".txt")

function Write-Report {
    param([string]$Text = "")
    Add-Content -LiteralPath $report -Value $Text -Encoding UTF8
}

Write-Report "=== DocumentExtractor Security Diagnosis ==="
Write-Report ("Time: " + (Get-Date))
Write-Report ("App: " + $AppPath)
Write-Report ""

if (!(Test-Path -LiteralPath $AppPath)) {
    Write-Report "RESULT: EXE missing"
    Write-Host ("Report: " + $report)
    exit 1
}

$item = Get-Item -LiteralPath $AppPath
Write-Report ("File: " + $item.FullName)
Write-Report ("Size: " + $item.Length)
Write-Report ("LastWriteTime: " + $item.LastWriteTime)

$hash = Get-FileHash -Algorithm SHA256 -LiteralPath $AppPath
Write-Report ("SHA256: " + $hash.Hash)

try {
    $bytes = [System.IO.File]::ReadAllBytes($AppPath)
    $headLength = [Math]::Min(16, $bytes.Length)
    $head = $bytes[0..($headLength - 1)]
    $hex = ($head | ForEach-Object { $_.ToString("X2") }) -join " "
    $asciiLength = [Math]::Min(8, $bytes.Length)
    $ascii = [System.Text.Encoding]::ASCII.GetString($bytes[0..($asciiLength - 1)])
    Write-Report ("HeaderHex: " + $hex)
    Write-Report ("HeaderAscii: " + $ascii)

    if ($ascii -eq "SCDSA004") {
        Write-Report "DIAGNOSIS: The EXE itself is wrapped by SoftCamp/DRM. Windows cannot run it as a PE executable."
    } elseif ($bytes[0] -eq 0x4D -and $bytes[1] -eq 0x5A) {
        Write-Report "DIAGNOSIS: EXE header is normal MZ/PE. If blocked, it is likely an execution policy/EDR/SmartScreen issue."
    } else {
        Write-Report "DIAGNOSIS: EXE header is neither MZ nor SCDSA004. File may be damaged or transformed."
    }
} catch {
    Write-Report ("Header check failed: " + $_.Exception.Message)
}

Write-Report ""
Write-Report "--- Authenticode ---"
try {
    Get-AuthenticodeSignature -LiteralPath $AppPath |
        Format-List * |
        Out-String |
        ForEach-Object { Write-Report $_ }
} catch {
    Write-Report ("Authenticode error: " + $_.Exception.Message)
}

Write-Report "--- Alternate Data Streams ---"
try {
    Get-Item -LiteralPath $AppPath -Stream * |
        Format-Table Stream,Length -AutoSize |
        Out-String |
        ForEach-Object { Write-Report $_ }
} catch {
    Write-Report ("Stream check error: " + $_.Exception.Message)
}

Write-Report "--- Defender Status ---"
try {
    Get-MpComputerStatus |
        Select-Object AMServiceEnabled,AntivirusEnabled,RealTimeProtectionEnabled,BehaviorMonitorEnabled,IoavProtectionEnabled,NISEnabled |
        Format-List |
        Out-String |
        ForEach-Object { Write-Report $_ }
} catch {
    Write-Report ("Defender status unavailable: " + $_.Exception.Message)
}

Write-Report "--- Defender Threats ---"
try {
    Get-MpThreat |
        Select-Object ThreatName,ActionSuccess,Resources,InitialDetectionTime,LastThreatStatusChangeTime |
        Format-List |
        Out-String |
        ForEach-Object { Write-Report $_ }
} catch {
    Write-Report ("Defender threat query unavailable: " + $_.Exception.Message)
}

Write-Report "--- Block Related Event Logs (last 3 days) ---"
$logs = @(
    "Microsoft-Windows-Windows Defender/Operational",
    "Microsoft-Windows-AppLocker/EXE and DLL",
    "Microsoft-Windows-CodeIntegrity/Operational"
)
$start = (Get-Date).AddDays(-3)
foreach ($log in $logs) {
    Write-Report ("### " + $log)
    try {
        if (Get-WinEvent -ListLog $log -ErrorAction SilentlyContinue) {
            Get-WinEvent -FilterHashtable @{ LogName = $log; StartTime = $start } -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.Message -match "DocumentExtractor|newppt|blocked|denied|deny|block|차단|SCDSA|SoftCamp|PyInstaller|$($hash.Hash)"
                } |
                Select-Object TimeCreated,Id,ProviderName,Message -First 30 |
                Format-List |
                Out-String |
                ForEach-Object { Write-Report $_ }
        }
    } catch {
        Write-Report ("Event query error: " + $_.Exception.Message)
    }
}

if ($RunStartTest -eq "1") {
    Write-Report "--- Start Test ---"
    try {
        $p = Start-Process -FilePath $AppPath -PassThru -ErrorAction Stop
        Start-Sleep -Seconds 5
        $alive = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
        if ($alive) {
            Write-Report ("Start test: process started, pid=" + $alive.Id)
            Stop-Process -Id $alive.Id -Force -ErrorAction SilentlyContinue
            Write-Report "Start test: process stopped after diagnosis"
        } else {
            Write-Report "Start test: process exited before 5 seconds"
        }
    } catch {
        Write-Report ("Start test failed: " + $_.Exception.Message)
    }
}

Write-Report ""
Write-Report "=== End ==="
Write-Host ("Report: " + $report)
