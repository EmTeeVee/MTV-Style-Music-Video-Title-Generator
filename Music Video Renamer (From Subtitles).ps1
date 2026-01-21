#---------------- CLEAR SCREEN ----------------
cls

Add-Type -AssemblyName System.Windows.Forms
$script:artist = $null
$script:title = $null
$script:album = $null
$script:year = $null

# ---------------- CHECKS ----------------
$requiredCommands = @('ffprobe', 'ffmpeg', 'mkvmerge')
foreach ($cmd in $requiredCommands) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        Write-Host "Error: $cmd not found in PATH."
        exit
    }
}

# ---------------- FOLDER SELECTION ----------------
$ownerForm = New-Object System.Windows.Forms.Form
$ownerForm.TopMost = $true
$ownerForm.StartPosition = 'Manual'
$ownerForm.Size = New-Object System.Drawing.Size(1,1)
$ownerForm.ShowInTaskbar = $false
$ownerForm.Opacity = 0
$ownerForm.Show()

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select folder containing videos with title cards:"

if ($folderBrowser.ShowDialog($ownerForm) -ne [System.Windows.Forms.DialogResult]::OK) {
    $ownerForm.Close()
    exit
}

$directory = $folderBrowser.SelectedPath
$ownerForm.Close()

# ---------------- SANITIZE FUNCTION ----------------
function Sanitize-Filename {
    param($s)
    if (-not $s) { return $null }
    # remove formatting codes first
    $s = $s -replace '\{.*?\}',''
    # replace non-Windows-compatible characters
    $s = $s -replace '/', '∕'
    $s = $s -replace ':', '꞉'
    $s = $s -replace '[<>\\|?*"]',''
    $s = $s -replace '\s{2,}',' '
    return $s.Trim()
}

# ---------------- VIDEO PROCESSING ----------------
$videoExtensions = @("*.mkv","*.mp4","*.avi","*.webm","*.mov")

Get-ChildItem -Path $directory -File | Where-Object {
    $_.Extension -match '\.(mkv|mp4|avi|webm|mov)$'
} | ForEach-Object {

    $video = $_
    $tempSrt = Join-Path $env:TEMP "temp_sub.srt"
    Remove-Item $tempSrt -ErrorAction SilentlyContinue

    # ---------- FIND FIRST SUBTITLE STREAM ----------
    $subInfo = & ffprobe -v error `
        -select_streams s `
        -show_entries stream=index,codec_name `
        -of csv=p=0 "$($video.FullName)" |
        Select-Object -First 1

    if (-not $subInfo) {
        Write-Host "❌ $($video.Name) — no subtitle stream found"
        return
    }

    $subIndex, $subCodec = $subInfo -split ','

    # ---------- EXTRACT SUBTITLE ----------
    & ffmpeg -y -loglevel error `
        -i "$($video.FullName)" `
        -map 0:$subIndex "$tempSrt"

    if (-not (Test-Path $tempSrt)) {
        Write-Host "❌ $($video.Name) — subtitle extraction failed"
        return
    }

    # ---------- READ FIRST SUBTITLE BLOCK ----------
    $lines = Get-Content $tempSrt | Where-Object { $_.Trim() -ne "" }

    if ($lines.Count -lt 6) {
        Write-Host "❌ $($video.Name) — less than 4 metadata lines, skipping"
        return
    }

    $artist = Sanitize-Filename $lines[2]
    $title  = Sanitize-Filename $lines[3]
    $album  = Sanitize-Filename $lines[4]
    $year   = Sanitize-Filename ($lines[5] -replace '[^\d]','')

    if (-not ($artist -and $title -and $album -and $year)) {
        Write-Host "❌ $($video.Name) — metadata parse failed"
        return
    }

    # ---------- BUILD NEW NAME ----------
    $newName = "$artist - $title ($year, $album)$($video.Extension)"
    $newPath = Join-Path $video.DirectoryName $newName

    if (-not (Test-Path $newPath)) {
        Rename-Item -LiteralPath $video.FullName -NewName $newName
        Write-Host "✅ Renamed → $newName"
    } else {
        Write-Host "⚠ Target already exists, skipping: $newName"
    }
}

Write-Host "`nDone."
