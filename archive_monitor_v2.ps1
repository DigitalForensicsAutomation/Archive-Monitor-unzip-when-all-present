#Requires -Version 5.1

# ================= CONFIG =================
$BasePath     = "C:\Archive"
$StableWait   = 5
$MaxRetries   = 3
$SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
# =========================================

$Incoming  = Join-Path $BasePath "Incoming"
$Temp      = Join-Path $BasePath "Temp"
$Completed = Join-Path $BasePath "Completed"
$Processed = Join-Path $BasePath "Processed"
$Failed    = Join-Path $BasePath "Failed"
$LogFile   = Join-Path $BasePath "archive.log"

# === FIND 7-ZIP ===
if (-not (Test-Path $SevenZipPath)) {
    # Try to find it via Get-Command if the hardcoded path fails
    $cmd = Get-Command "7z.exe" -ErrorAction SilentlyContinue
    if ($cmd) { 
        $SevenZipPath = $cmd.Source 
    } else {
        $SevenZipPath = "7z" # Fallback
    }
}

# === ENSURE DIRECTORIES ===
foreach ($dir in @($Incoming,$Temp,$Completed,$Processed,$Failed)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

# === LOGGING ===
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message

    switch ($Level) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }

    try {
        Add-Content -Path $LogFile -Value $line -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to write to log file: $_" -ForegroundColor Red
    }
}

# === AUDIO ALERTS ===
function Beep-NewFile { [console]::Beep(1200,150); Start-Sleep -Milliseconds 100; [console]::Beep(1200,150) }
function Beep-ExtractStart { [console]::Beep(600,200); Start-Sleep -Milliseconds 50; [console]::Beep(900,200) }
function Beep-ExtractSuccess { [console]::Beep(1400,400) }
function Beep-ExtractFail { [console]::Beep(400,500) }

# === FILE STABILITY CHECK ===
function Wait-FileStable {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return $false }

    # Initial wait to let write start
    Start-Sleep -Seconds 2
    
    $size1 = (Get-Item $Path).Length
    Start-Sleep -Seconds $StableWait

    if (-not (Test-Path $Path)) { return $false }

    $size2 = (Get-Item $Path).Length
    return ($size1 -eq $size2)
}

# === GET SPLIT PARTS ===
function Get-ArchiveParts {
    param([string]$FirstPart)
    $dir  = Split-Path $FirstPart
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FirstPart)
    Get-ChildItem -Path $dir -File | Where-Object { $_.Name -match "^$base\.\d{3}$" } | Sort-Object Name
}

# === VERIFY PARTS ===
function Test-PartsComplete {
    param($Parts)
    if ($Parts.Count -lt 2) { return $false }
    for ($i = 0; $i -lt $Parts.Count; $i++) {
        $expected = "{0:D3}" -f ($i + 1)
        if ($Parts[$i].Name -notmatch "\.$expected$") { return $false }
    }
    return $true
}

# === MOVE PARTS ===
function Move-Parts {
    param($Parts, [string]$Destination)
    foreach ($p in $Parts) {
        if (Test-Path $p.FullName) {
            Move-Item -LiteralPath $p.FullName -Destination (Join-Path $Destination $p.Name) -Force -ErrorAction SilentlyContinue
        }
    }
}

# === EXTRACT ARCHIVE ===
function Extract-Archive {
    param([string]$Archive, [string]$OutDir)
    $proc = Start-Process -FilePath $SevenZipPath -ArgumentList "x `"$Archive`" -o`"$OutDir`" -y" -Wait -NoNewWindow -PassThru
    return ($proc.ExitCode -le 1)
}

# === PROCESS SINGLE ARCHIVE ===
# Refactored to handle a specific file object passed by the watcher
function Process-ArchiveFile {
    param($ArchiveFileItem)

    $a = $ArchiveFileItem
    Write-Log "Detected archive $($a.Name)"
    Beep-NewFile

    # Wait for stability
    if (-not (Wait-FileStable $a.FullName)) {
        Write-Log "File $($a.Name) is unstable or vanished. Skipping." "WARN"
        return
    }

    $parts = Get-ArchiveParts $a.FullName
    if (-not (Test-PartsComplete $parts)) {
        Write-Log "Archive $($a.Name) incomplete (parts missing). Waiting." "WARN"
        return
    }

    $workDir = Join-Path $Temp $a.BaseName
    if (Test-Path $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $workDir -Force | Out-Null

    $success = $false

    for ($i = 1; $i -le $MaxRetries; $i++) {
        Write-Log "Extraction attempt $i for $($a.Name)"
        Beep-ExtractStart

        if (Extract-Archive $a.FullName $workDir) {
            if ((Get-ChildItem $workDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0) {
                $success = $true
                break
            }
        }
        Start-Sleep -Seconds 3
    }

    if ($success) {
        Write-Log "Extraction successful for $($a.Name)" "SUCCESS"
        Beep-ExtractSuccess
        
        $archiveRoot = Join-Path $Completed $a.BaseName
        $files = Get-ChildItem -Path $workDir -Recurse -File -ErrorAction SilentlyContinue
        
        foreach ($file in $files) {
            if (-not (Test-Path $file.FullName)) { continue }
            $rel  = $file.FullName.Substring($workDir.Length).TrimStart('\')
            $dest = Join-Path $archiveRoot $rel
            $dir  = Split-Path $dest
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Move-Item -LiteralPath $file.FullName -Destination $dest -Force -ErrorAction SilentlyContinue
        }

        Start-Sleep -Seconds 2
        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
        Move-Parts $parts $Processed
    }
    else {
        Write-Log "Extraction failed for $($a.Name)" "ERROR"
        Beep-ExtractFail
        Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
        Move-Parts $parts $Failed
    }
    Write-Log "-----------------------------------"
}

# === START MONITOR (FileSystemWatcher) ===
Write-Log "Archive Monitor Started (Event Driven)" "SUCCESS"
Write-Log "Watching $Incoming for *.001 files..."

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $Incoming
$watcher.Filter = "*.001"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

# Define the action to take when a file is created
$action = { 
    $path = $Event.SourceEventArgs.FullPath
    $name = $Event.SourceEventArgs.Name
    $changeType = $Event.SourceEventArgs.ChangeType
    
    # Small sleep to ensure file handle is released by the system creation event
    Start-Sleep -Milliseconds 500
    
    # Convert path to Item object to match original script logic
    if (Test-Path $path) {
        $item = Get-Item $path
        Process-ArchiveFile -ArchiveFileItem $item
    }
}

# Register the event
# We use separate registrations for Created to capture new files
Register-ObjectEvent $watcher "Created" -Action $action | Out-Null

# Keep the script running indefinitely to listen for events
try {
    while ($true) {
        Start-Sleep -Seconds 5
        # Perform a manual check periodically in case an event was missed
        # (Optional redundancy)
        $missed = Get-ChildItem -Path $Incoming -Filter "*.001" -File
        foreach ($m in $missed) {
             # Only process if older than 10 seconds to avoid conflict with Watcher
             if ($m.LastWriteTime -lt (Get-Date).AddSeconds(-10)) {
                 Process-ArchiveFile -ArchiveFileItem $m
             }
        }
    }
}
finally {
    # Clean up
    Unregister-Event -SourceIdentifier "FileSystemWatcher.Created" -ErrorAction SilentlyContinue
    $watcher.Dispose()
    Write-Log "Monitor Stopped"
}
