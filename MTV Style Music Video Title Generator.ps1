#---------------- WARNING ----------------
cls
Write-Host "⚠️ WARNING! Destructive script. ⚠️ Overwrites video files."

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Continue"

Add-Type -AssemblyName System.Windows.Forms
$script:knownTags = @('SD', '4K', 'AV1', 'VP9', 'VOB')

# list to track duplicates and failed files
$duplicates = @()
$failedFiles = @()

# ---------------- REQUIRED PROGRAMS ----------------
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

# ---------------- MUSIC VIDEO FILENAME CHECK ----------------
$videoExtensions = '.mkv', '.mp4', '.avi', '.vob', '.webm', '.ts', '.mov', '.wmv', '.m4v', '.flv'
$videoFiles = Get-ChildItem -Path $directory -File | Where-Object { $videoExtensions -contains $_.Extension.ToLower() }

$invalidFiles = @()

foreach ($file in $videoFiles) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

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
        try { Rename-Item -LiteralPath $_.FullName -NewName $NewName } catch {Write-Warning "Failed processing: $filePath"}
    }
}

# ---------------- VIDEO DURATION ----------------
function Get-VideoDuration {
    param($filePath)
    if (-not (Test-Path -LiteralPath $filePath)) { return 0 }
    try {
        $args = @('-v','error','-show_entries','format=duration','-of','default=noprint_wrappers=1:nokey=1',"$filePath")
        [math]::Round([double](& ffprobe @args))
    } catch { 0 }
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

# ---------------- VIDEO RESOLUTION ----------------
function Get-VideoResolution {
    param($filePath)
    $global:playResX = 0
    $global:playResY = 0

    if (-not $filePath -or -not (Test-Path -LiteralPath $filePath)) {
        Write-Host "WARNING: Invalid file path for ffprobe: '$filePath'"
        return
    }

    try {
        $args = @('-v','error','-select_streams','v:0','-show_entries','stream=width,height','-of','csv=p=0:s=x',"$filePath")
        $output = (& ffprobe @args).Trim()
        $match = [regex]::Match($output, '(\d+)\s*[x,]\s*(\d+)')
        if ($match.Success) {
            $global:playResX = [int]$match.Groups[1].Value
            $global:playResY = [int]$match.Groups[2].Value
            return
        }

        # fallback method
        $heightStr = (& ffprobe -v error -select_streams v:0 -show_entries stream=height -of csv=p=0 "$filePath" | Select-Object -First 1)
        $widthStr  = (& ffprobe -v error -select_streams v:0 -show_entries stream=width  -of csv=p=0 "$filePath" | Select-Object -First 1)

        if ([int]::TryParse($widthStr.Trim(), [ref]$global:playResX) -and [int]::TryParse($heightStr.Trim(), [ref]$global:playResY)) { return }

        Write-Host "WARNING: Unable to detect resolution for '$filePath'"

    } catch {
        Write-Host "WARNING: ffprobe failed for '$filePath'"
        $global:playResX = 0
        $global:playResY = 0
    }
}

# ---------------- AUDIO CHANNELS ----------------
function Get-AudioChannelLayouts {
    param([string]$videoPath)
    $channels = @()
    try {
        $streamInfo = & ffprobe -v error -select_streams a -show_entries stream=channels -of csv=p=0 "$videoPath"
        foreach ($line in $streamInfo) {
            if ([int]::TryParse($line, [ref]$null)) {
                $channels += [int]$line
            }
        }
    } catch {Write-Warning "Failed processing: $filePath"}
    return $channels
}

# ---------------- LYRIC LANGUAGE ----------------
function Get-LanguageFromSuffix {
    param($name)
    if ($name -match '_([a-z]{3})\.(srt|ass)$') {
        return $matches[1]
    }
    return 'und'
}

# ---------------- CREATE SUBTITLES ----------------
function Create-ASS {
    param($filePath, $duration)

    $artist  = $null
    $title   = $null
    $year    = $null
    $album   = $null
    $orgartist = $null
    $extras  = @()
    $version = $null

    $assFile  = [System.IO.Path]::ChangeExtension($filePath,'.ass')
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)

    # ---------------- PARSE FILENAME ----------------
    $parseName = $fileName -replace '\s*\{[^}]+\}\s*$',''
    $parseName = $parseName -replace '[∕]', '/'
    $parseName = $parseName -replace '[꞉]', ':'
    $parseName = $parseName.Trim()
    $lines = @()

    if (-not ($parseName -match '^(?<artist>.+?)\s-\s(?<rest>.+)$')) { return }

    $artist = $matches.artist.Trim()
    $rest = $matches.rest.Trim()

    $yearAlbumMatch = [regex]::Match(
        $rest,
        '\((?<year>\d{4})(?:,\s*(?<album>(?>[^\(\)]|\((?<d>)|\)(?<-d>))*(?(d)(?!))))?\)'
    )

    if ($yearAlbumMatch.Success) {
        $year = $yearAlbumMatch.Groups['year'].Value
        $album = if ($yearAlbumMatch.Groups['album'].Success) { $yearAlbumMatch.Groups['album'].Value.Trim() } else { $null }

        $titlePart = $rest.Substring(0, $yearAlbumMatch.Index).Trim()

        $orgartist = $null
        $preYearBrackets = [regex]::Matches($titlePart, '\[[^\]]+\]')
        if ($preYearBrackets.Count -gt 0) {
            $orgartist = "[$($preYearBrackets[0].Value.Trim('[', ']')) Cover]"
            $titlePart = $titlePart -replace [regex]::Escape($preYearBrackets[0].Value), ''
        }

        $title = $titlePart.Trim()

        $afterYear = $rest.Substring($yearAlbumMatch.Index + $yearAlbumMatch.Length).Trim()
        $extras = @()
        foreach ($m in [regex]::Matches($afterYear, '\[[^\]]+\]')) {
            $extras += $m.Value
        }
    } else {
        $title = $rest
        $orgartist = $null
        $album = $null
        $extras = @()
    }

    # ---------------- TITLE LINE WITH QUOTES ----------------
    $extras = $extras | Where-Object {
        $script:knownTags -notcontains $_.ToUpper()
    }

    $titleText = $title
    if ($version) { $titleText += ' ' + $version }

    # opening and closing quotes with more space to the right and left respectively [modified Kabel-Black.ttf needed]
    # $titleLineFinal = '"' + $titleText + '"' # use this if original Kabel-Black.ttf is used
    $oq     = [char]0xE000
    $cq     = [char]0xE001
    $oqnegative  = [char]0xE002
    $cqnegative  = [char]0xE003
    $oqlarge  = [char]0xE004
    $cqlarge  = [char]0xE005

    $openQuote  = $oq
    $closeQuote = $cq

    if (![string]::IsNullOrEmpty($titleText)) {
            
    $firstChar = $titleText[0]
    if ($firstChar -cin @('A','J')) {
        $openQuote = $oqnegative
    }
    if ($firstChar -cin @('T','V','W','X','Y')) {
        $openQuote = $oqlarge
    }

    $lastChar = $titleText[$titleText.Length - 1]
    if ($lastChar -cin @('A','L')) {
    $closeQuote = $cqnegative
    }
    if ($lastChar -cin @('T','V','W','X','Y')) {
    $closeQuote = $cqlarge
    }
    }

$titleLineFinal = $openQuote + $titleText + $closeQuote
    
    if (@($extras).Count -gt 0) {
        $titleLineFinal += ' ' + (@($extras) -join ' ')
    }

    # ---------------- BUILD LINES ----------------
    function Clean-AssText { param($text) if (-not $text) { return $null }; return ($text -replace "(`r`n|`n|`r)", "") }

    $lines = @()
    if ($artist)   { $lines += Clean-AssText $artist }
    if ($titleLineFinal)  { $lines += Clean-AssText $titleLineFinal }
    if ($orgartist)       { $lines += Clean-AssText $orgartist }
    if ($album)           { $lines += "{\i1}$(Clean-AssText $album){\i0}" }
    if ($year)     { $lines += Clean-AssText $year }

    # ---------------- RESOLUTION ----------------
    $playResX = $global:playResX
    $playResY = $global:playResY
    if ($playResX -le 0 -or $playResY -le 0) {
        Write-Warning "Invalid resolution for '$filePath', skipping ASS creation"
        return
        }

    # ---------------- FONT SCALING ----------------
    $baseFontSize    = 66
    $baseLineSpacing = 53
    $lineCount = 4
    $gapCount  = $lineCount - 1
    $referenceBlockHeight = ($lineCount * $baseFontSize) + ($gapCount * $baseLineSpacing)

    # effective display height for 16:9 screen
    $referenceHeight = 1080.0
    $referenceAspect = 16.0 / 9.0
    $videoAspect = $playResX / $playResY

    if ($videoAspect -gt $referenceAspect) {
        $effectiveHeight = $playResX / $referenceAspect
    }
    else {
        $effectiveHeight = $playResY
    }

    $heightScale = $effectiveHeight / $referenceHeight
    $targetBlockHeight = $referenceBlockHeight * $heightScale

    $spacingRatio = $baseLineSpacing / $baseFontSize
    $floatFontSize = $targetBlockHeight / ($lineCount + ($gapCount * $spacingRatio))
    $fontSize = [math]::Round($floatFontSize)
    $lineSpacing = [math]::Round($spacingRatio * $fontSize)
   
    # ---------------- SHADOW / OUTLINE FIX ----------------
    $areaScaleFactor = [math]::Sqrt(($playResX * $playResY) / (1920 * 1080))
    $outline = [math]::Round(2 * $areaScaleFactor)
    $shadow  = [math]::Ceiling(7 * $areaScaleFactor)
    if ($outline -lt 1) { $outline = 1 }
    if ($shadow -le 2) { $shadow = 2 }
   
    # ---------------- TITLE CARD ----------------
    $ass = @()
    $ass += "[Script Info]"
    $ass += "Title: Title Card"
    $ass += "ScriptType: v4.00+"
    $ass += "PlayResX: $playResX"
    $ass += "PlayResY: $playResY"
    $ass += "Collisions: Normal"
    $ass += "WrapStyle: 2"
    $ass += "ScaledBorderAndShadow: yes"
    $ass += ""
    $ass += "[V4+ Styles]"
    $ass += "Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding"
    $ass += "Style: Default,Kabel-Black,$fontSize,&H00FFFFFF,&H00FFFFFF,&H00000000,&H64000000,0,0,0,0,100,100,0,0,1,$outline,$shadow,1,75,0,63,1"
    $ass += ""
    $ass += "[Events]"
    $ass += "Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text"

    # ---------------- TIMINGS ----------------
    $introStart = 3
    $introEnd   = $introStart + 13
    $outroStart = [math]::Max($duration - 16, 0)
    $outroEnd   = [math]::Max($duration - 3, 0)

    # ---------------- POSITIONING ----------------
    $lineCount = $lines.Count

    # DPI-independent 2.5cm approximate margins
    $cmToPixels = 96 / 2.54
    $marginCm = 2.5

    # use the smaller scale factor (width or height) for consistent physical margins
    $scaleFactor = [math]::Min($playResX / 1920, $playResY / 1080)
    $leftMargin   = [math]::Round($marginCm * $cmToPixels * $scaleFactor)
    $bottomMargin = [math]::Round($marginCm * $cmToPixels * $scaleFactor)

    # total height of the block
    $blockHeight = ($lineCount - 1) * $lineSpacing

    # bottom-alignment
    $blockBottomY = $playResY - $bottomMargin
    $topY = $blockBottomY - $blockHeight
    $scaledOffset = [math]::Round(10 * ($playResY / 1080))

    for ($i = 0; $i -lt $lineCount; $i++) {
        $x = $leftMargin
        $y = [math]::Round($topY + ($i * $lineSpacing))+$scaledOffset
 
        $ass += "Dialogue: 0,$(SecondsToASSTime $introStart),$(SecondsToASSTime $introEnd),Default,,0,0,0,," +
                "{\pos($x,$y)\alpha&HFF&\t(0,5000,2.5,\alpha&H00&)\t(8000,13000,2.5,\alpha&HFF&)}$($lines[$i])"

        $ass += "Dialogue: 0,$(SecondsToASSTime $outroStart),$(SecondsToASSTime $outroEnd),Default,,0,0,0,," +
                "{\pos($x,$y)\alpha&HFF&\t(0,5000,2.5,\alpha&H00&)\t(8000,13000,2.5,\alpha&HFF&)}$($lines[$i])"
    }

    # ---------------- SAVE ----------------
    $ass -join "`r`n" | Set-Content -LiteralPath $assFile -Encoding UTF8

    return [PSCustomObject]@{
    AssPath = $assFile
    Artist  = $artist
    Title   = $title
    Year    = $year
}
}

#---------------- PROCESS LYRIC FILE AND CONVERT TO ASS ----------------
function Process-LyricFile {
    param([string]$lyricPath)
    
    # read *.srt lines
    $lines = Get-Content -LiteralPath $lyricPath -Encoding UTF8
    $processed = @()
    $capitalizeNext = $true

    # ---------- capitalization & text cleanup ----------
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

        # detect lines that are mostly uppercase
        if ($trim -match '[A-Z]' -and $trim -cmatch '^[^a-z]*[A-Z]{2,}') {
            $line = $line.ToLower()
            if ($capitalizeNext) {
                $line = [regex]::Replace($line,'^(\W*)([a-z])',{ $args[0].Groups[1].Value + $args[0].Groups[2].Value.ToUpper() })
            }
            # replace special characters
            $line = $line -replace "[´``’�]", "'"
            $line = $line -replace "[´``‘’]", "'"
            $line = $line -replace "[“”]", '"'
            $line = $line -replace "[–—−]", '-'
            $line = $line -replace "…", "..."

            # fix common pronouns
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

    # ---------- convert to *.ass ----------
    $ass = @"
[Script Info]
ScriptType: v4.00+
PlayResX: 1920
PlayResY: 1080

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,Arial,42,&H00FFFFFF,&H000000FF,&H00000000,&H64000000,0,0,0,0,100,100,0,0,1,2,0,2,60,60,40,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"@ -split "`n"

$dialogues = @()
$prevEnd = [TimeSpan]::Zero
$currentText = @()
$currentStart = $null
$currentEnd = $null

foreach ($line in $processed) {
    $trim = $line.Trim()

    if ($trim -match '^\d+$') { continue }

    if ($trim -match '-->') {
        # finish previous dialogue
        if ($currentStart -ne $null -and $currentText.Count -gt 0) {
            if ($prevEnd -gt [TimeSpan]::Zero -and $currentStart -lt $prevEnd) {
                # truncate previous end to avoid overlap
                $currentEnd = $currentStart
            }
            $text = ($currentText -join '\N')
            $dialogues += "Dialogue: 0,{0},{1},Default,,0,0,0,,{2}" -f `
                $currentStart.ToString("h\:mm\:ss\.ff"), `
                $currentEnd.ToString("h\:mm\:ss\.ff"), `
                $text
            $prevEnd = $currentEnd
            $currentText = @()
        }

        $times = $trim -split ' --> '
        $currentStart = [TimeSpan]::ParseExact($times[0], "hh\:mm\:ss\,fff", $null)
        $currentEnd   = [TimeSpan]::ParseExact($times[1], "hh\:mm\:ss\,fff", $null)
        continue
    }

    if ($trim -eq '') { continue }

    $currentText += $trim
}

# add last subtitle
if ($currentStart -ne $null -and $currentText.Count -gt 0) {
    if ($prevEnd -gt [TimeSpan]::Zero -and $currentStart -lt $prevEnd) {
        $currentEnd = $currentStart
    }
    $text = ($currentText -join '\N')
    $dialogues += "Dialogue: 0,{0},{1},Default,,0,0,0,,{2}" -f `
        $currentStart.ToString("h\:mm\:ss\.ff"), `
        $currentEnd.ToString("h\:mm\:ss\.ff"), `
        $text
}

$ass += $dialogues

# write *.ass file
$assPath = [System.IO.Path]::ChangeExtension($lyricPath, ".ass")
$ass | Set-Content -LiteralPath $assPath -Encoding UTF8

return $assPath
}

# ---------------- MUX SUBTITLES ----------------
function Mux-SubtitlesAndFinalize {
    param(
        [string]$videoPath,
        [string]$titleAss,
        [string]$artist,
        [string]$title,
        [string]$year
    )

$dir  = Split-Path -LiteralPath $videoPath
$base = [System.IO.Path]::GetFileNameWithoutExtension($videoPath)

# ---------------- collect lyric subtitles ----------------
$allSRTs = @(Get-ChildItem -LiteralPath $dir -File -Filter "*.srt" -ErrorAction SilentlyContinue)
$lyricsAll = @($allSRTs | Where-Object { $_.Name -match "^$([regex]::Escape($base))_" })

$lyricsProcessed = @()
foreach ($l in $lyricsAll) {
    $lyricsProcessed += Process-LyricFile $l.FullName
}

# ---------------- temp output ----------------
$tempMKV = Join-Path -Path $dir -ChildPath "__WORKFILE__.mkv"

# ---------------- audio analysis ----------------
$channels = Get-AudioChannelLayouts -videoPath $videoPath
if (-not $channels) { $channels = @() }
elseif (-not ($channels -is [System.Collections.IEnumerable])) { $channels = @($channels) }

# ---------------- existing subtitles ----------------
$json = & mkvmerge -J "$videoPath" | ConvertFrom-Json
$subtitleTracks = @($json.tracks | Where-Object { $_.type -eq "subtitles" })
$excludeSubtitleTrackIds = @()

# ---------------- external lyric languages ----------------
$newLyricsLanguages = @()
if (@($lyricsProcessed).Count -gt 0) {
    $newLyricsLanguages = @(
        $lyricsProcessed |
        ForEach-Object { Get-LanguageFromSuffix $_ } |
        Where-Object { $_ } |
        Select-Object -Unique
    )
}

foreach ($track in $subtitleTracks) {
    $trackName = $track.properties.track_name
    $lang      = $track.properties.language
    $id        = $track.id

    $keep = $false
    if ($trackName -eq "Lyrics" -and $lang -and $lang -notin $newLyricsLanguages) {
        $keep = $true
    }

    if (-not $keep) {
        $excludeSubtitleTrackIds += $id
    }
}

# ---------------- build mkvmerge arguments ----------------
$args = @("--quiet", "--output", "$tempMKV")

# ---- TAGGING ----
if ($artist -and $title) {
    $args += "--title"
    $args += "$artist • $title"
}
else {
    $args += "--title"
    $args += "$base"
}

if ($year) {
    # --- segment date ---
    $args += "--date"
    $args += "$year-06-06T00:00:00Z"

 # --- Global Plex-friendly tags for zero-process muxing ---
$tagFile = Join-Path $dir "__plex_tags.xml"

$tagXml = @"
<?xml version="1.0"?>
<Tags>
  <Tag>
    <Targets />
    <Simple>
      <Name>DATE_RELEASED</Name>
      <String>$year</String>
    </Simple>
    <Simple>
      <Name>DATE</Name>
      <String>$year</String>
    </Simple>
    <Simple>
      <Name>FILE_CREATED_DATE</Name>
      <String>$year</String>
    </Simple>
    <Simple>
      <Name>FILE_MODIFIED_DATE</Name>
      <String>$year</String>
    </Simple>
    <Simple>
      <Name>ORIGINAL_RELEASE_DATE</Name>
      <String>$year-06-06</String>
    </Simple>
    <Simple>
      <Name>YEAR</Name>
      <String>$year</String>
    </Simple>
    <Simple>
      <Name>ARTIST</Name>
      <String>$artist</String>
    </Simple>
    <Simple>
      <Name>TITLE</Name>
      <String>$title</String>
    </Simple>
  </Tag>
</Tags>
"@

# write the *.xml to a temporary file
$tagXml | Set-Content -LiteralPath $tagFile -Encoding UTF8

# add global tags
$args += "--global-tags"
$args += "$tagFile"
}

if ($titleAss -and (Test-Path -LiteralPath $titleAss)) {
    $args += "--language";      $args += "0:eng"
    $args += "--track-name";    $args += "0:Title Card"
    $args += "--default-track"; $args += "0:no"
    $args += "--forced-track";  $args += "0:yes"
    $args += $titleAss
}

for ($i = 0; $i -lt @($channels).Count; $i++) {
    $label = switch ($channels[$i]) {
        1 { "Mono" }
        2 { "Stereo" }
        6 { "Surround 5.1" }
        8 { "Surround 7.1" }
        default { "Surround $($channels[$i])" }
    }
    $trackId = $i + 1
    $args += "--language";   $args += "${trackId}:eng"
    $args += "--track-name"; $args += "${trackId}:$label"
}

if (@($excludeSubtitleTrackIds).Count -gt 0) {
    $args += "--subtitle-tracks"
    $args += ("!" + ($excludeSubtitleTrackIds -join ","))
}
$args += "$videoPath"

for ($i = 0; $i -lt @($lyricsProcessed).Count; $i++) {
    $lang = Get-LanguageFromSuffix ([System.IO.Path]::GetFileName($lyricsProcessed[$i]))
    $args += "--language";   $args += "0:$lang"
    $args += "--track-name"; $args += "0:Lyrics"
    $args += "--default-track"; $args += "0:" + ($(if ($lang -eq "eng") { "yes" } else { "no" }))
    $args += $lyricsProcessed[$i]
}

# ---------------- attachment handling ----------------
$attachments = @($json.attachments)
if ($fontPath -and (Test-Path -LiteralPath $fontPath)) {
    foreach ($att in $attachments) {
        & mkvpropedit "$videoPath" --delete-attachment $att.id 1>$null
    }
    $args += "--attachment-mime-type"; $args += "application/x-truetype-font"
    $args += "--attach-file";          $args += "$fontPath"
}

Write-Host "Muxing $($base).mkv …"
& mkvmerge @args 1>$null
$mkvExitCode = $LASTEXITCODE

if ($mkvExitCode -gt 1 -or -not (Test-Path $tempMKV)) {
    throw "mkvmerge failed"
}

# ---------------- finalize ----------------
$final = Join-Path -Path $dir -ChildPath ($base + ".mkv")
if (Test-Path -LiteralPath $final) {
    Remove-Item -LiteralPath $final -Force
}
Rename-Item -LiteralPath $tempMKV -NewName $final

# ---------------- filesystem timestamps (for Plex) ----------------
if ($year) {
    $newDate = Get-Date "$year-06-06T00:00:00Z"

    $fileInfo = Get-Item -LiteralPath $final
    $fileInfo.CreationTime   = $newDate
    $fileInfo.LastWriteTime  = $newDate
    $fileInfo.LastAccessTime = $newDate
}

# ---------------- VIDEO TAGGER {SD}, {4K}, {AV1}, {VP9}, {VOB}----------------

$file = Get-Item -LiteralPath $final
$baseName = $file.BaseName
$ext      = $file.Extension
$tags     = @()

$probeJson = ffprobe -v error -select_streams v:0 `
    -show_entries stream=width,height,codec_name `
    -print_format json "$file" | ConvertFrom-Json

if ($probeJson.streams) {
    $width  = [int]$probeJson.streams[0].width
    $height = [int]$probeJson.streams[0].height
    $codec  = $probeJson.streams[0].codec_name.ToLower()

    $sdHeightMin = 720
    $fourKHeight = 2160
    $fourKWidth  = 3840
    $fourKHeightRatio = 0.95
    $fourKWidthRatio  = 0.95

    $dvdResolutions = @(
        @{Width=720; Height=480},
        @{Width=720; Height=576}
    )

    $isDVD = $false
    foreach ($res in $dvdResolutions) {
        if ($width -eq $res.Width -and $height -eq $res.Height) {
            $isDVD = $true
            break
        }
    }

    if ($height -lt $sdHeightMin -or $isDVD) { $tags += "SD" }

    $fourKHeightThreshold = [math]::Round($fourKHeight * $fourKHeightRatio)
    $fourKWidthThreshold  = [math]::Round($fourKWidth  * $fourKWidthRatio)

    if ($height -ge $fourKHeightThreshold -or $width -ge $fourKWidthThreshold) {
        $tags += "4K"
    }

    foreach ($uc in "AV1","VP9","VOB") {
    if ($codec -like "*$uc*") { $tags += $uc }
}

# ---------------- extract existing wavy brackets {...} and build new ones ----------------
if ($tags.Count -gt 0) {
    $existingBracketMatch = [regex]::Match($baseName, '\{([^}]+)\}\s*$')

    $existingCustomTags   = @()
    $existingAnalysisTags = @()

    if ($existingBracketMatch.Success) {

        $existingParts = $existingBracketMatch.Groups[1].Value -split '\s*,\s*'

        foreach ($part in $existingParts) {

            $trimmed = $part.Trim()

            if ($trimmed -in @("SD","4K","AV1","VP9","VOB")) {
                $existingAnalysisTags += $trimmed
            }
            elseif ($trimmed) {
                $existingCustomTags += $trimmed
            }
        }
    }

    $cleanBase = ($baseName -replace '\s*\{[^}]+\}\s*$', '').Trim()

    if ($existingCustomTags -contains "Logo") {

        $withoutLogo = @()
        foreach ($tag in $existingCustomTags) {
            if ($tag -ne "Logo") {
                $withoutLogo += $tag
            }
        }

        if ($withoutLogo.Count -gt 1) {
            $newCustom = @()
            $newCustom += $withoutLogo[0]
            $newCustom += "Logo"

            for ($i = 1; $i -lt $withoutLogo.Count; $i++) {
                $newCustom += $withoutLogo[$i]
            }

            $existingCustomTags = $newCustom
        }
        elseif ($withoutLogo.Count -eq 1) {
            $existingCustomTags = @($withoutLogo[0], "Logo")
        }
        else {
            $existingCustomTags = @("Logo")
        }
    }

    $finalTags = @()

    if ($existingCustomTags.Count -gt 0) {
        $finalTags += $existingCustomTags
    }

    $allAnalysis = @()
    $allAnalysis += $existingAnalysisTags
    $allAnalysis += $tags

    if ($allAnalysis.Count -gt 0) {
        $allAnalysis = $allAnalysis | Sort-Object -Unique
        $finalTags  += $allAnalysis
    }

    if ($finalTags.Count -eq 0) {
        return
    }

    $newName = $cleanBase + " {" + ($finalTags -join ", ") + "}" + $ext
    $dir     = [System.IO.Path]::GetDirectoryName($file.FullName)
    $newPath = [System.IO.Path]::Combine($dir, $newName)

    try {
        [System.IO.File]::Move($file.FullName, $newPath)
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}
}

# ---------------- cleanup ----------------
if ([System.IO.Path]::GetExtension($videoPath).ToLower() -ne ".mkv") {
    Remove-Item -LiteralPath $videoPath -Force -ErrorAction SilentlyContinue
}
if ($titleAss -and (Test-Path -LiteralPath $titleAss)) {
    Remove-Item -LiteralPath $titleAss -Force -ErrorAction SilentlyContinue
}
if ($tagFile -and (Test-Path -LiteralPath $tagFile)) {
    Remove-Item -LiteralPath $tagFile -Force -ErrorAction SilentlyContinue
}
foreach ($l in @($lyricsAll))       { Remove-Item -LiteralPath $l.FullName -Force -ErrorAction SilentlyContinue }
foreach ($t in @($lyricsProcessed)) { Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue }
}

# ---------------- PROCESS VIDEOS ----------------
$videoFiles = Get-ChildItem $directory -Recurse -File -Include *.mp4,*.mkv,*.avi,*.vob,*.webm,*.ts,*.mov,*.wmv,*.m4v,*.flv
foreach ($file in $videoFiles) {
    try {
        $dur = Get-VideoDuration $file.FullName
        if ($dur -le 0) {
            Write-Warning "Skipping (invalid duration): $($file.Name)"
            continue
        }

        Get-VideoResolution $file.FullName

        try {
        $assResult = Create-ASS $file.FullName $dur
        }
        catch {
            Write-Warning "ASS creation failed: $($file.Name)"
            continue
        }

        Mux-SubtitlesAndFinalize `
            $file.FullName `
            $assResult.AssPath `
            $assResult.Artist `
            $assResult.Title `
            $assResult.Year
    }
   catch {
    $failedFiles += [PSCustomObject]@{
        File  = $file.Name
        Stage = "Processing"
        Error = $_.Exception.Message
    }
    
    Write-Warning "Failed processing: $($file.Name)"
    continue
}
}

# ------------------ SHOW DUPLICATES ------------------
if (@($duplicates).Count -gt 0) {
    Write-Host "`nFiles renamed with numbering due to duplicates:"
    $duplicates | Format-Table -AutoSize
}

# ------------------ SHOW FAILURES ------------------
if (@($failedFiles).Count -gt 0) {
    Write-Host "`n⚠️ Files that failed to process:"
    $failedFiles | Format-Table -AutoSize
}

Write-Host "All files processed."
Read-Host "Press Enter to exit"
