#Requires -Version 5.1
<#
.SYNOPSIS
    Installs a staged XLSmart add-in after Excel releases the active file.

.DESCRIPTION
    This helper is launched by the add-in before Excel exits. It waits for the
    calling Excel process (and any process holding the add-in) to stop, verifies
    the staged payload, creates a timestamped backup, replaces the add-in
    atomically, and rolls back on any failure.

    Failed updates preserve the staged payload and backup for diagnosis.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$StagedPath,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Fa-f0-9]{64}$')]
    [string]$ExpectedSha256,

    [Parameter(Mandatory = $false)]
    [int]$ExcelPid = 0,

    [Parameter(Mandatory = $false)]
    [ValidateRange(10, 600)]
    [int]$TimeoutSeconds = 120,

    [Parameter(Mandatory = $false)]
    [string]$StatusPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:StatusPath = $StatusPath
$script:BackupPath = $null
$script:InstallTempPath = $null
$script:OriginalExists = $false
$script:ReplacedDestination = $false

function Write-Status {
    param([Parameter(Mandatory = $true)][string]$Message)

    $line = "{0:u} {1}" -f (Get-Date), $Message
    Write-Host $line

    if (-not [string]::IsNullOrWhiteSpace($script:StatusPath)) {
        try {
            $statusDirectory = Split-Path -Parent $script:StatusPath
            if ($statusDirectory -and -not (Test-Path -LiteralPath $statusDirectory)) {
                New-Item -ItemType Directory -Path $statusDirectory -Force | Out-Null
            }
            Add-Content -LiteralPath $script:StatusPath -Value $line -Encoding UTF8
        }
        catch {
            Write-Host "Status log unavailable: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Get-Hash {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Test-FileAvailable {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $true
    }

    $handle = $null
    try {
        $handle = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($handle) { $handle.Dispose() }
    }
}

function Wait-ForExcelExit {
    param(
        [Parameter(Mandatory = $false)][int]$PidToWait = 0,
        [Parameter(Mandatory = $true)][int]$Seconds = 120
    )

    $deadline = (Get-Date).AddSeconds($Seconds)

    # Wait for the originating Excel process. Do not kill it or any other instance.
    if ($PidToWait -gt 0) {
        $callingProc = Get-Process -Id $PidToWait -ErrorAction SilentlyContinue
        if ($callingProc) {
            Write-Status "Waiting for calling Excel process (PID $PidToWait) to exit..."
            while ($callingProc -and -not $callingProc.HasExited -and (Get-Date) -lt $deadline) {
                Start-Sleep -Milliseconds 500
                try { $callingProc.Refresh() } catch { $callingProc = $null }
                if ($callingProc -and $callingProc.HasExited) { $callingProc = $null }
            }
        }
    }

    # Any other Excel instance may also hold the target add-in. Abort rather than
    # waiting on unrelated sessions indefinitely or replacing a loaded file.
    $otherExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($PidToWait -gt 0) {
        $otherExcel = @($otherExcel | Where-Object { $_.Id -ne $PidToWait })
    }
    if ($otherExcel.Count -gt 0) {
        Write-Status 'Another Excel instance is still running; refusing to replace the add-in.'
        return $false
    }

    # PID exit alone does not guarantee the OS has released the add-in file.
    do {
        if (Test-FileAvailable -Path $DestinationPath) {
            Write-Status 'Destination file lock released.'
            return $true
        }
        Start-Sleep -Seconds 1
    } while ((Get-Date) -lt $deadline)

    return (Test-FileAvailable -Path $DestinationPath)
}

function Restore-Backup {
    if (-not $script:OriginalExists) {
        if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
            $failedPath = "$DestinationPath.failed.$([Guid]::NewGuid().ToString('N').Substring(0,8))"
            Move-Item -LiteralPath $DestinationPath -Destination $failedPath -Force
            Write-Status "Quarantined failed installation: $failedPath"
        }
        return
    }

    if (-not (Test-Path -LiteralPath $script:BackupPath -PathType Leaf)) {
        throw "Backup file missing: $script:BackupPath"
    }

    $restoreTemp = "$DestinationPath.rollback.$([Guid]::NewGuid().ToString('N')).tmp"
    Copy-Item -LiteralPath $script:BackupPath -Destination $restoreTemp -Force

    if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
        $failedPath = "$DestinationPath.failed.$([Guid]::NewGuid().ToString('N').Substring(0,8))"
        [System.IO.File]::Replace($restoreTemp, $DestinationPath, $failedPath, $true)
        Write-Status "Preserved failed installation at: $failedPath"
    }
    else {
        [System.IO.File]::Move($restoreTemp, $DestinationPath)
    }

    $restoredHash = Get-Hash $DestinationPath
    $originalHash = Get-Hash $script:BackupPath
    if ($restoredHash -ne $originalHash) {
        throw "Rollback hash mismatch. Expected $originalHash, got $restoredHash."
    }
}

try {
    $ExpectedSha256 = $ExpectedSha256.ToUpperInvariant()
    $destinationDirectory = Split-Path -Parent $DestinationPath
    if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
    }

    Write-Status 'XLSmart update helper started.'
    $stagedHash = Get-Hash $StagedPath
    if ($stagedHash -ne $ExpectedSha256) {
        throw "Staged file hash mismatch. Expected $ExpectedSha256, got $stagedHash."
    }
    Write-Status "Staged payload hash verified: $stagedHash"

    if (-not (Wait-ForExcelExit -PidToWait $ExcelPid -Seconds $TimeoutSeconds)) {
        throw "Excel or file lock did not release within $TimeoutSeconds seconds. Close Excel and retry."
    }

    $script:OriginalExists = Test-Path -LiteralPath $DestinationPath -PathType Leaf
    if ($script:OriginalExists) {
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $guidShort = [Guid]::NewGuid().ToString('N').Substring(0, 8)
        $script:BackupPath = "$DestinationPath.backup_${stamp}_$guidShort"
        Copy-Item -LiteralPath $DestinationPath -Destination $script:BackupPath -Force
        Write-Status "Backup created: $script:BackupPath"
    }

    $script:InstallTempPath = "$DestinationPath.update.$([Guid]::NewGuid().ToString('N')).tmp"
    Copy-Item -LiteralPath $StagedPath -Destination $script:InstallTempPath -Force
    if ((Get-Hash $script:InstallTempPath) -ne $ExpectedSha256) {
        throw 'Temporary install copy failed hash verification.'
    }

    if ($script:OriginalExists) {
        [System.IO.File]::Replace($script:InstallTempPath, $DestinationPath, $script:BackupPath, $true)
        Write-Status "Atomic replacement completed: $DestinationPath"
    }
    else {
        [System.IO.File]::Move($script:InstallTempPath, $DestinationPath)
        Write-Status "New add-in installed: $DestinationPath"
    }
    $script:ReplacedDestination = $true
    $script:InstallTempPath = $null

    $installedHash = Get-Hash $DestinationPath
    if ($installedHash -ne $ExpectedSha256) {
        throw "Installed file hash mismatch. Expected $ExpectedSha256, got $installedHash."
    }

    Remove-Item -LiteralPath $StagedPath -Force -ErrorAction SilentlyContinue
    Write-Status "XLSmart update installed successfully: $DestinationPath"
    if ($script:BackupPath) {
        Write-Status "Previous version retained at: $script:BackupPath"
    }
    exit 0
}
catch {
    $failure = $_.Exception.Message
    Write-Status "Update failed: $failure"

    try {
        if ($script:InstallTempPath -and (Test-Path -LiteralPath $script:InstallTempPath -PathType Leaf)) {
            Remove-Item -LiteralPath $script:InstallTempPath -Force -ErrorAction SilentlyContinue
        }
        if ($script:ReplacedDestination) {
            Write-Status 'Attempting rollback from backup...'
            Restore-Backup
            Write-Status 'Rollback completed and verified.'
        }
    }
    catch {
        Write-Status "Rollback failed: $($_.Exception.Message)"
        Write-Status "Keep backup and staged files for manual recovery."
    }

    exit 1
}
