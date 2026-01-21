#---------------- CLEAR SCREEN ----------------
cls
Write-Host "⚠️ WARNING! Destructive script. ⚠️ Overwrites video files."

Add-Type -AssemblyName System.Windows.Forms

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
$folderBrowser.Description = "⚠️ Files will be overwritten! Select folder containing videos:"

if ($folderBrowser.ShowDialog($ownerForm) -ne [System.Windows.Forms.DialogResult]::OK) {
    $ownerForm.Close()
    exit
}

$directory = $folderBrowser.SelectedPath
$ownerForm.Close()

# ---------------- FONT CHECK ----------------
$fontFileName = "Kabel-Black.ttf"
$fontPath = Join-Path $directory $fontFileName

if (-not (Test-Path -LiteralPath $fontPath)) {
    # Font not found in the selected folder, prompt user to select it
    $fontForm = New-Object System.Windows.Forms.Form
    $fontForm.TopMost = $true
    $fontForm.StartPosition = 'CenterScreen'
    $fontForm.Size = New-Object System.Drawing.Size(1,1)
    $fontForm.ShowInTaskbar = $false
    $fontForm.Opacity = 0
    $fontForm.Show()

    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Kabel-Black.ttf not found in selected folder. Please locate the font file."
    $fileDialog.Filter = "TrueType Font (*.ttf)|*.ttf"

    if ($fileDialog.ShowDialog($fontForm) -eq [System.Windows.Forms.DialogResult]::OK) {
        $fontPath = $fileDialog.FileName
    } else {
        Write-Host "Font file not selected. Exiting."
        $fontForm.Close()
        exit
    }

    $fontForm.Close()
}

# ---------------- MUSIC VIDEO FILENAME SANITY CHECK ----------------
# Only include video files
$videoExtensions = '.mkv', '.mp4', '.avi', '.vob', '.webm', '.ts', '.mov', '.wmv', '.m4v', '.flv'
$videoFiles = Get-ChildItem -Path $directory -File | Where-Object { $videoExtensions -contains $_.Extension.ToLower() }

$invalidFiles = @()

foreach ($file in $videoFiles) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    # Strict 'artist - title' check
    if ($name -notmatch '^\s*\S.*\S\s-\s\S.*\S\s*$') {
        $invalidFiles += $file.Name
    }
}

if ($invalidFiles.Count -gt 0) {
    Write-Host "⚠️ The following files do not match the expected 'artist - title' format:"
    $invalidFiles | ForEach-Object { Write-Host " - $_" }
    Write-Host "`nAre you sure this is the correct folder? These may not be music video files."
    exit
}

# ---------------- NORMALIZE FILENAMES ----------------
Get-ChildItem -Path $directory -File | ForEach-Object {
    $OldName = $_.Name
    $NewName = $OldName

    $fancyApostrophes = @([char]0x2018, [char]0x2019, [char]0x201B, [char]0x00B4, [char]0x0060, [char]0x2032)
    foreach ($c in $fancyApostrophes) {$NewName = $NewName.Replace($c, "'")}

    $NewName = $NewName.Replace([char]0x2013,"-")
    $NewName = $NewName.Replace([char]0x2014,"-")
    $NewName = $NewName.Replace([char]0x2015,"-")

    while ($NewName -like "*  *") { $NewName = $NewName -replace '  ', ' ' }
    $NewName = $NewName.Replace("…", "...")
    $NewName = $NewName.Trim()

    if ($NewName -ne $OldName) {
        try { Rename-Item -LiteralPath $_.FullName -NewName $NewName } catch {}
    }
}

# ---------------- VIDEO DURATION ----------------
function Get-VideoDuration {
    param($filePath)
    try { [math]::Round([double](& ffprobe -v error -show_entries format=duration `
        -of default=noprint_wrappers=1:nokey=1 "$filePath")) } catch { 0 }
}

function SecondsToASSTime {
    param([double]$seconds)
    $totalSeconds = [math]::Max(0, [math]::Round($seconds,2))
    $hours   = [int][math]::Floor($totalSeconds / 3600)
    $minutes = [int][math]::Floor(($totalSeconds % 3600) / 60)
    $secs    = [int][math]::Floor($totalSeconds % 60)
    $centis  = [int][math]::Floor(($totalSeconds - [math]::Floor($totalSeconds)) * 100)
    return "{0}:{1:D2}:{2:D2}.{3:D2}" -f $hours, $minutes, $secs, $centis
}

# ---------------- AUDIO CHANNELS ----------------
function Get-AudioChannelLayouts {
    param([string]$videoPath)
    $channels = @()
    try {
        $streamInfo = & ffprobe -v error -select_streams a `
            -show_entries stream=channels -of csv=p=0 "$videoPath"
        foreach ($line in $streamInfo) {
            if ([int]::TryParse($line, [ref]$null)) {
                $channels += [int]$line
            }
        }
    } catch {}
    return $channels
}

# ---------------- LYRIC LANGUAGE ----------------
function Get-LanguageFromSuffix {
    param($name)
    if ($name -match '_([a-z]{3})\.srt$') { $matches[1] } else { 'und' }
}

# ---------------- CREATE ASS ----------------
function Create-ASS {
    param($filePath, $duration)

    # ---- RESET PER FILE ----
    $script:artist = $null
    $script:title  = $null
    $script:year   = $null
    $album         = $null
    $orgartist     = $null
    $extras        = @()

    $assFile  = [System.IO.Path]::ChangeExtension($filePath,'.ass')
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)

    # ---------------- PARSE FILENAME ----------------
    $parseName = $fileName -replace '\s*\{[^}]+\}\s*$',''
    $parseName = $parseName -replace '[∕]', '/'
    $parseName = $parseName -replace '[꞉]', ':'
    $parseName = $parseName.Trim()
    $lines = @()

    if (-not ($parseName -match '^(?<artist>.+?)\s-\s(?<rest>.+)$')) { return }

    $script:artist = $matches.artist.Trim()
    $rest = $matches.rest.Trim()

    $yearAlbumMatch = [regex]::Match(
        $rest,
        '\((?<year>\d{4})(?:,\s*(?<album>(?>[^\(\)]|\((?<d>)|\)(?<-d>))*(?(d)(?!))))?\)'
    )

    if ($yearAlbumMatch.Success) {
        $script:year = $yearAlbumMatch.Groups['year'].Value
        $album = if ($yearAlbumMatch.Groups['album'].Success) { $yearAlbumMatch.Groups['album'].Value.Trim() } else { $null }

        $titlePart = $rest.Substring(0, $yearAlbumMatch.Index).Trim()

        $orgartist = $null
        $preYearBrackets = [regex]::Matches($titlePart, '\[[^\]]+\]')
        if ($preYearBrackets.Count -gt 0) {
            $orgartist = "[$($preYearBrackets[0].Value.Trim('[', ']')) Cover]"
            $titlePart = $titlePart -replace [regex]::Escape($preYearBrackets[0].Value), ''
        }

        $script:title = $titlePart.Trim()

        $afterYear = $rest.Substring($yearAlbumMatch.Index + $yearAlbumMatch.Length).Trim()
        $extras = @()
        foreach ($m in [regex]::Matches($afterYear, '\[[^\]]+\]')) {
            $extras += $m.Value
        }
    } else {
        $script:title = $rest
        $orgartist = $null
        $album = $null
        $extras = @()
    }

    # ---------------- TITLE LINE WITH QUOTES ----------------
    $negSpace = '{\fsp2}'
    $reset    = '{\fsp0}'

    $titleText = $script:title
    if ($version) { $titleText += ' ' + $version }

    $titleLineFinal = '"' + $negSpace + $titleText + $reset + '"'
    if ($extras.Count -gt 0) { $titleLineFinal += ' ' + ($extras -join ' ') }

    # ---------------- BUILD LINES ----------------
    function Clean-AssText { param($text) if (-not $text) { return $null }; return ($text -replace "(`r`n|`n|`r)", "") }

    $lines = @()
    if ($script:artist)   { $lines += Clean-AssText $script:artist }
    if ($titleLineFinal)  { $lines += Clean-AssText $titleLineFinal }
    if ($orgartist)       { $lines += Clean-AssText $orgartist }
    if ($album)           { $lines += "{\i1}$(Clean-AssText $album){\i0}" }
    if ($script:year)     { $lines += Clean-AssText $script:year }

    # ---------------- ASS HEADER ----------------
    $ass = @()
    $ass += "[Script Info]"
    $ass += "Title: Title Card"
    $ass += "ScriptType: v4.00+"
    $ass += "PlayResX: 1920"
    $ass += "PlayResY: 1080"
    $ass += "Collisions: Normal"
    $ass += "WrapStyle: 2"
    $ass += "ScaledBorderAndShadow: yes"
    $ass += ""
    $ass += "[V4+ Styles]"
    $ass += "Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding"
    $ass += "Style: Default,Kabel-Black,66,&H00FFFFFF,&H00FFFFFF,&H00000000,&H64000000,0,0,0,0,100,100,0,0,1,1,4,1,75,0,63,1"
    $ass += ""
    $ass += "[Events]"
    $ass += "Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text"

    # ---------------- TIMINGS (INTRO + OUTRO) ----------------
    $introStart = 3
    $introEnd   = $introStart + 13
    $outroStart = [math]::Max($duration - 16, 0)
    $outroEnd   = [math]::Max($duration - 3, 0)
    if ($introEnd -le $introStart) { $introEnd = $introStart + 0.01 }
    if ($outroEnd -le $outroStart) { $outroEnd = $outroStart + 0.01 }

    # ---------------- POSITIONING ----------------
    $playResX = 1920
    $playResY = 1080
    $leftFraction   = 0.04    # horizontal position
    $bottomFraction = 0.80    # vertical position of first line in original script
    $lineSpacing    = 55      # vertical spacing

    $lineCount = $lines.Count

    # ---------------- FIXED BOTTOM LINE BASED ON 4-LINE REFERENCE ----------------
    $bottomY4 = [math]::Round($playResY * $bottomFraction + (3 * $lineSpacing))
    $topY     = $bottomY4 - (($lineCount - 1) * $lineSpacing) # shift top line so bottom aligns

    for ($i = 0; $i -lt $lineCount; $i++) {
        $x = [math]::Round($playResX * $leftFraction)
        $y = [math]::Round($topY + ($i * $lineSpacing))

        # Intro line
        $ass += "Dialogue: 0,$(SecondsToASSTime $introStart),$(SecondsToASSTime $introEnd),Default,,0,0,0,," +
                "{\pos($x,$y)\alpha&HFF&\t(0,5000,2.5,\alpha&H00&)\t(8000,13000,2.5,\alpha&HFF&)}$($lines[$i])"

        # Outro line
        $ass += "Dialogue: 0,$(SecondsToASSTime $outroStart),$(SecondsToASSTime $outroEnd),Default,,0,0,0,," +
                "{\pos($x,$y)\alpha&HFF&\t(0,5000,2.5,\alpha&H00&)\t(8000,13000,2.5,\alpha&HFF&)}$($lines[$i])"
    }

    # ---------------- SAVE ----------------
    $ass -join "`r`n" | Set-Content -LiteralPath $assFile -Encoding UTF8

    return $assFile
}

#---------------- PROCESS LYRIC FILE ----------------
function Process-LyricFile {
    param([string]$lyricPath)
    
    $lines = Get-Content -LiteralPath $lyricPath -Encoding UTF8
    $processed = @()
    $capitalizeNext = $true

    foreach ($line in $lines) {
        $trim = $line.Trim()

        if ($trim -match '^\d+$' -or $trim -match '-->') {
            $processed += $line
            $capitalizeNext = $true
            continue
        }
        if ($trim -eq '') {
            $processed += $line
            continue
        }
        if ($trim -match '^\(.*\)$') {
            $processed += $line.ToLower()
            $capitalizeNext = $true
            continue
        }

        if ($trim -match '[A-Z]' -and $trim -cmatch '^[^a-z]*[A-Z]{2,}') {
            $line = $line.ToLower()
            if ($capitalizeNext) {
                $line = [regex]::Replace($line,'^(\W*)([a-z])',{ $args[0].Groups[1].Value + $args[0].Groups[2].Value.ToUpper() })
            }
            $line = $line -replace "[´``’�]", "'"
            $line = $line -replace "[´``‘’]", "'"
            $line = $line -replace "[“”]", '"'
            $line = $line -replace "[–—−]", '-'
            $line = $line -replace "…", "..."
            $line = $line -replace '\bi\b','I'
            $line = $line -replace '\bi''m\b',"I'm"
            $line = $line -replace '\bi''ve\b',"I've"
            $line = $line -replace '\bi''ll\b',"I'll"
            $line = $line -replace '\bi''d\b',"I'd"
            $line = $line -replace '([.!?]\s+)i\b', '$1I'
            $capitalizeNext = $false
        }

        $processed += $line
    }

    $tempPath = Join-Path ([System.IO.Path]::GetDirectoryName($lyricPath)) ("__temp_" + [System.IO.Path]::GetFileName($lyricPath))
    $processed | Set-Content -LiteralPath $tempPath -Encoding UTF8
    return $tempPath
}

# ---------------- MUXING ----------------
function Mux-SubtitlesAndFinalize {
    param(
        [string]$videoPath,
        [string]$titleAss
    )

    $dir  = Split-Path $videoPath
    $base = [System.IO.Path]::GetFileNameWithoutExtension($videoPath)

    # ---------------- Collect lyric subtitles ----------------
    $allSRTs = Get-ChildItem $dir -File -Filter "*.srt"
    $lyricsAll = $allSRTs | Where-Object {
        $_.Name -match "^$([regex]::Escape($base))_"
    }

    $lyricsProcessed = @()
    foreach ($l in $lyricsAll) {
        $lyricsProcessed += Process-LyricFile $l.FullName
    }

    # ---------------- Temp output ----------------
    $tempMKV = Join-Path $dir "__WORKFILE__.mkv"

    # ---------------- Audio analysis ----------------
    $channels = Get-AudioChannelLayouts $videoPath

   # ---------------- Existing subtitles in the video file ----------------
$json = & mkvmerge -J "$videoPath" | ConvertFrom-Json
$subtitleTracks = $json.tracks | Where-Object { $_.type -eq "subtitles" }
$excludeSubtitleTrackIds = @()

# Determine which lyric languages exist externally for the file
$newLyricsLanguages = @(
    $lyricsProcessed |
    ForEach-Object { Get-LanguageFromSuffix $_ } |
    Where-Object { $_ }
) | Select-Object -Unique

foreach ($track in $subtitleTracks) {

    $trackName = $track.properties.track_name
    $lang      = $track.properties.language
    $codec     = $track.codec
    $id        = $track.id

    $keep = $false

    # ---- Embedded Lyrics logic ----
    if ($trackName -eq "Lyrics" -and $lang) {

        # Keep only if no new external lyric exists for this language
        if ($lang -notin $newLyricsLanguages) {
            $keep = $true
        }
    }

    # ---- All other subtitles are removed ----
    if (-not $keep) {
        $excludeSubtitleTrackIds += $id
    }
}


    # ---------------- Build mkvmerge arguments ----------------
    $args = @(
        "--quiet",
        "--output", $tempMKV,
        "--title",  $base
    )

    # ---------- Title ASS (forced, first subtitle) ----------
    $args += "--language";       $args += "0:eng"
    $args += "--track-name";     $args += "0:Title Card"
    $args += "--default-track";  $args += "0:yes"
    $args += "--forced-track";   $args += "0:yes"
    $args += $titleAss

    # ---------- Audio: language + track names ----------
    for ($i = 0; $i -lt $channels.Count; $i++) {
        switch ($channels[$i]) {
            1 { $label = "Mono" }
            2 { $label = "Stereo" }
            6 { $label = "Surround 5.1" }
            8 { $label = "Surround 7.1" }
            default { $label = "Surround $($channels[$i])" }
        }

        $trackId = $i + 1
        $args += "--language"; $args += "${trackId}:und"
        $args += "--track-name"; $args += "${trackId}:$label"
    }

    # ---------- Main video file ----------
    if ($excludeSubtitleTrackIds.Count -gt 0) {
        $args += "--subtitle-tracks"
        $args += ("!" + ($excludeSubtitleTrackIds -join ","))
    }
    $args += $videoPath

    # ---------- Lyric subtitles ----------
    for ($i = 0; $i -lt $lyricsProcessed.Count; $i++) {
        $lang = Get-LanguageFromSuffix ([System.IO.Path]::GetFileName($lyricsProcessed[$i]))
        $args += "--language";   $args += "0:$lang"
        $args += "--track-name"; $args += "0:Lyrics"
        if ($lang -eq "eng") { $args += "--default-track"; $args += "0:no" }
        $args += $lyricsProcessed[$i]
    }

    # ---------- Font attachment (Kabel-Black.ttf) ----------
    if ($fontPath -and (Test-Path -LiteralPath $fontPath)) {
        # Remove all existing attachments from the original video
        if ($json.attachments) {
            foreach ($att in $json.attachments) {
                & mkvpropedit "$videoPath" --delete-attachment $att.id 1>$null
            }
        }
        $args += "--attachment-mime-type"; $args += "application/x-truetype-font"
        $args += "--attach-file"; $args += $fontPath
    }

    # ---------------- Run mkvmerge ----------------
    Write-Host "Muxing $($base).mkv …"
    & mkvmerge @args 1>$null

    if ($LASTEXITCODE -ne 0) {
        Write-Warning "⚠️ mkvmerge failed for $videoPath"
        return
    }

    # ---------------- Finalize ----------------
    $final = Join-Path $dir ($base + ".mkv")
    if (Test-Path -LiteralPath $final) { Remove-Item -LiteralPath $final -Force }
    Rename-Item -LiteralPath $tempMKV -NewName $final

    # ---------------- Plex-visible Metadata ----------------
    if ($script:year) {
        $dateString = "$script:year-05-13T00:00:00Z"
        & mkvpropedit "$final" --edit info --set "date=$dateString" 1>$null
    }
    if ($script:artist -and $script:title) {
        $plexTitle = "$script:artist • $script:title" # this is not visible in 'Other Videos' category 
        & mkvpropedit "$final" --edit info --set "title=$plexTitle" 1>$null
    }

    # ---------------- Cleanup ----------------
    if ([System.IO.Path]::GetExtension($videoPath).ToLower() -ne ".mkv") {
        Remove-Item -LiteralPath $videoPath -Force -ErrorAction SilentlyContinue
    }
    if ($titleAss -and (Test-Path -LiteralPath $titleAss)) {
        Remove-Item -LiteralPath $titleAss -Force -ErrorAction SilentlyContinue
    }
    foreach ($l in $lyricsAll)       { Remove-Item -LiteralPath $l.FullName -Force -ErrorAction SilentlyContinue }
    foreach ($t in $lyricsProcessed) { Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue }
}

# ---------------- PROCESS VIDEOS ----------------
$videoFiles = Get-ChildItem $directory -Recurse -File -Include *.mp4,*.mkv,*.avi,*.vob,*.webm,*.ts,*.mov,*.wmv,*.m4v,*.flv
foreach ($file in $videoFiles) {
    $dur = Get-VideoDuration $file.FullName
    if ($dur -gt 0) {
        $titleAss = Create-ASS $file.FullName $dur
        Mux-SubtitlesAndFinalize $file.FullName $titleAss
    }
}

# ---------------- VIDEO TAGGER {SD}, {4K}, {AV1}, {VP9}, {VOB}----------------

# Supported video extensions
$videoExtensions = "*.mp4","*.mkv","*.avi","*.vob","*.webm","*.ts","*.mov","*.wmv","*.m4v","*.flv"

# List to track duplicates
$duplicates = @()

# ---------------- CONFIG ----------------
$sdHeightMin      = 720
$fourKHeight      = 2160
$fourKWidth       = 3840
$fourKHeightRatio = 0.95
$fourKWidthRatio  = 0.95

# DVD resolutions
$dvdResolutions = @(
    @{Width=720; Height=480},  # NTSC
    @{Width=720; Height=576}   # PAL
)

# Unusual codecs
$unusualCodecs = "av1","vp9","vob"

# ---------------- SCRIPT ----------------
Get-ChildItem -Path $directory -Recurse -File -Include $videoExtensions | ForEach-Object {

    $file = $_.FullName
    $baseName = $_.BaseName
    $ext = $_.Extension
    $tags = @()

    # SAFEGUARD: skip files already tagged with any of these
    if ($baseName -match '\{(SD|4K|av1|vp9|vob)(\s*\d+)?\}') { return }

    # ------------------ GET VIDEO INFO ------------------
    $heightStr = ffprobe -v error -select_streams v -show_entries stream=height -of csv=p=0 "$file" | Select-Object -First 1
    $widthStr  = ffprobe -v error -select_streams v -show_entries stream=width  -of csv=p=0 "$file" | Select-Object -First 1
    $codecStr  = ffprobe -v error -select_streams v -show_entries stream=codec_name -of csv=p=0 "$file" | Select-Object -First 1

    if (-not $heightStr -or -not $widthStr -or -not $codecStr) { return }

    # Safe trimming
    $heightStr = $heightStr.Trim()
    $widthStr  = $widthStr.Trim()
    $codecStr  = $codecStr.Trim().ToLower()

    # Convert height/width to int
    [int]$height = 0
    [int]$width  = 0
    if (-not [int]::TryParse($heightStr,[ref]$height)) { return }
    if (-not [int]::TryParse($widthStr ,[ref]$width )) { return }

    $codec = $codecStr

    # ------------------ SD DETECTION ------------------
    $isDVD = $false
    foreach ($res in $dvdResolutions) {
        if ($width -eq $res.Width -and $height -eq $res.Height) { $isDVD = $true; break }
    }

    if ($height -lt $sdHeightMin -or $isDVD) { $tags += "SD" }

    # ------------------ 4K DETECTION ------------------
    $fourKHeightThreshold = [math]::Round($fourKHeight * $fourKHeightRatio)
    $fourKWidthThreshold  = [math]::Round($fourKWidth  * $fourKWidthRatio)

    if ($height -ge $fourKHeightThreshold -or $width -ge $fourKWidthThreshold) { $tags += "4K" }

    # ------------------ CODEC DETECTION ------------------
    foreach ($uc in $unusualCodecs) {
        if ($codec -like "*$uc*") { $tags += $uc }
    }

    # ------------------ REMOVE DUPLICATE TAGS ------------------
    $tags = $tags | Select-Object -Unique

    # ------------------ RENAME FILE ------------------
    if ($tags.Count -gt 0) {
        $tagString = "{"+($tags.ToUpper() -join ", ")+"}"

        # Avoid adding tag if already present
        if ($baseName -notmatch [regex]::Escape($tagString)) {

            $newName = "$baseName $tagString$ext"
            $newPath = Join-Path $_.DirectoryName $newName

            # Handle filename duplicates
            $counter = 0
            while (Test-Path $newPath) {
                $counter++
                $newName = "$baseName $tagString $counter$ext"
                $newPath = Join-Path $_.DirectoryName $newName
            }

            if ($counter -gt 0) {
                $duplicates += [PSCustomObject]@{
                    Original = $_.Name
                    Renamed  = $newName
                }
            }

            Rename-Item -LiteralPath $file -NewName $newName
        }
    }
}

# ------------------ SHOW DUPLICATES ------------------
if ($duplicates.Count -gt 0) {
    Write-Host "`nFiles renamed with numbering due to duplicates:"
    $duplicates | Format-Table -AutoSize
}

Write-Host "All files processed."