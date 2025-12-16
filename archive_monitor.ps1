#Requires -Version 5.1

# ================= CONFIG =================
$BasePath     = "C:\Archive"
$PollInterval = 30
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
    $SevenZipPath = "7z"
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

    Add-Content -Path $LogFile -Value $line
}

# === FILE STABILITY CHECK ===
function Wait-FileStable {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return $false }

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

    Get-ChildItem -Path $dir -File |
        Where-Object { $_.Name -match "^$base\.\d{3}$" } |
        Sort-Object Name
}

# === VERIFY PARTS ===
function Test-PartsComplete {
    param($Parts)

    if ($Parts.Count -lt 2) { return $false }

    for ($i = 0; $i -lt $Parts.Count; $i++) {
        $expected = "{0:D3}" -f ($i + 1)
        if ($Parts[$i].Name -notmatch "\.$expected$") {
            return $false
        }
    }
    return $true
}

# === MOVE PARTS ===
function Move-Parts {
    param($Parts, [string]$Destination)

    foreach ($p in $Parts) {
        if (Test-Path $p.FullName) {
            Move-Item -LiteralPath $p.FullName `
                      -Destination (Join-Path $Destination $p.Name) `
                      -Force -ErrorAction SilentlyContinue
        }
    }
}

# === EXTRACT ARCHIVE ===
function Extract-Archive {
    param([string]$Archive, [string]$OutDir)

    $proc = Start-Process `
        -FilePath $SevenZipPath `
        -ArgumentList "x `"$Archive`" -o`"$OutDir`" -y" `
        -Wait -NoNewWindow -PassThru

    return ($proc.ExitCode -le 1)
}

# === PROCESS LOOP ===
function Process-Archives {

    $archives = Get-ChildItem -Path $Incoming -Filter "*.001" -File |
                Sort-Object LastWriteTime

    foreach ($a in $archives) {

        Write-Log "Processing $($a.Name)"

        if (-not (Wait-FileStable $a.FullName)) {
            Write-Log "File still being written" "WARN"
            continue
        }

        $parts = Get-ArchiveParts $a.FullName
        if (-not (Test-PartsComplete $parts)) {
            Write-Log "Archive incomplete" "WARN"
            continue
        }

        # === PER-ARCHIVE WORK DIR ===
        $workDir = Join-Path $Temp $a.BaseName
        if (Test-Path $workDir) {
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null

        # === EXTRACT WITH RETRIES ===
        $success = $false
        for ($i = 1; $i -le $MaxRetries; $i++) {
            Write-Log "Extraction attempt $i"
            if (Extract-Archive $a.FullName $workDir) {
                if ((Get-ChildItem $workDir -Recurse -File -ErrorAction SilentlyContinue).Count -gt 0) {
                    $success = $true
                    break
                }
            }
            Start-Sleep -Seconds 3
        }

        if ($success) {

            Write-Log "Extraction successful" "SUCCESS"

            $archiveRoot = Join-Path $Completed $a.BaseName

            # === SNAPSHOT FILE LIST (CRITICAL FIX) ===
            $files = Get-ChildItem -Path $workDir -Recurse -File -ErrorAction SilentlyContinue

            foreach ($file in $files) {

                if (-not (Test-Path $file.FullName)) { continue }

                $rel  = $file.FullName.Substring($workDir.Length).TrimStart('\')
                $dest = Join-Path $archiveRoot $rel
                $dir  = Split-Path $dest

                if (-not (Test-Path $dir)) {
                    New-Item -ItemType Directory -Path $dir -Force | Out-Null
                }

                if (Test-Path $dest) {
                    $dest = Join-Path $dir "$(Get-Date -Format yyyyMMddHHmmss)_$($file.Name)"
                }

                Move-Item -LiteralPath $file.FullName `
                          -Destination $dest `
                          -Force -ErrorAction SilentlyContinue
            }

            Start-Sleep -Seconds 2
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            Move-Parts $parts $Processed
        }
        else {
            Write-Log "Extraction failed" "ERROR"
            Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
            Move-Parts $parts $Failed
        }

        Write-Log "-----------------------------------"
    }
}

# === START MONITOR ===
Write-Log "Archive Monitor Started" "SUCCESS"
Write-Log "Watching $Incoming"

while ($true) {
    try {
        Process-Archives
    }
    catch {
        Write-Log $_.Exception.Message "ERROR"
    }
    Start-Sleep -Seconds $PollInterval
}
