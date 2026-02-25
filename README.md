# MTV-Style-Music-Video-Title-Generator
![Screenshot](https://raw.githubusercontent.com/EmTeeVee/MTV-Style-Music-Video-Title-Generator/refs/heads/main/EmTeeVee.jpg)

Creates MTV Style Title Cards (*.ass subtitles) for music videos using their filenames
-------------------------------------------------------------------------
Please be aware that this scripts deletes/muxes all your original files,
so I strongly advise you to use it in a copied folder and
be extremely careful using it.
-------------------------------------------------------------------------
If you don't know what you are doing:  
DON'T DO IT!
-------------------------------------------------------------------------
For everyone:
Music Video Renamer (From Subtitles).ps1 is not part of this project.  
DO NOT USE IT!
-------------------------------------------------------------------------
For GP:
Music Video Renamer (From Subtitles).ps1 is to rename your files.
-------------------------------------------------------------------------
Installation requirements (programs should be in windows path):  
powershell (obviously)  
ffprobe  
ffmpeg  
mkvmerge  
altered MTV style font "Kabel-Black.ttf" should be available (this file contains a workaround with additional quotation mark characters that have improved LSB/RSB relative to the text as kerning does not work in *.ass files), ideally in the folder you are muxing in (the script will prompt you for it, if it is not there)

I built this so I can play music videos with title cards in Plex (all tag dates and also the file creation and modification time are now set to the middle of the year the video was released, as Plex keeps reading data differently during the initial database update and metadata refreshes later on - but, hopefully, sorting by year or decade should not be a problem anymore).

Please be aware that because of the way Plex works, it usually transcodes the video to show the *.ass subtitles contained in the final *.mkv files. If you use an NVidia shield or the Plex Desktop App it should work without transcoding, in other scenarios the transcoding of higher resolutions might cause playback problems.

Music videos must be named like this:  
"artist - title (remix) [original artist] (year, album) [extras] {genre}.videoextension"  
-------------------------------------------------------------------------
" - " has to be between artist and title (artist names containing a dash have to be written without spaces)
(year) or (year, album) is the second important distinguisher, everything else is optional  
at least "artist - title.ext" has to be present in every single file in the chosen directory (and its subdirectories), otherwise the program will not run
characters that can be used in filenames that are recognized/converted by the script are ∕ and ꞉

Examples:  
ABBA - I Can Be That Woman (2021, Voyage) [Lyric Video].mp4  
a-ha - Take On Me (1984, Hunting High And Low) [35mm].mp4
Billie Eilish - No Time To Die (2020, No Time To Die꞉ Soundtrack).mp4
Blondie - The Tide Is High [The Paragons] (1980, Autoamerican).mkv  
Siouxsie And The Banshees - Kiss Them For Me (1991, Superstition).vob
Tarja - Frosty The Snowman (2023, Dark Christmas) {Xmas}.mkv

Additionally, the script can add subtitle tracks to the music videos if you have them present alongside your music video file and named accordingly, for example:  
Eurythmics - Here Comes The Rain Again (1983, Touch).mp4  
Eurythmics - Here Comes The Rain Again (1983, Touch)_eng.srt  
Eurythmics - Here Comes The Rain Again (1983, Touch)_fre.srt  

Finally, some information is written at the end of the filename in wavy brackets {} after muxing, so that {SD}, {4K}, {AV1} or {VP9} files are marked. This is for quality and compatibilty reasons that one might have to consider.

I found it helpful to add {Xmas} at the end for christmas music, so that I can ask Plex to only play those tracks during hot summer nights...

Usage in Plex:
-------------------------------------------------------------------------
For the title cards to be visible in Plex, the music videos have to be added as 'other videos' (agent: personal media) and under settings - languages - subtitle mode "shown with foreign audio" and prefer subtitles in "english" has to be selected. If there are lyric subtitles as well inside the *.mkv files, those would be showing if the subtitle mode is set to "always enabled" - in this case, videos that do not have lyric subtitles will still show the *.ass title cards.

<img src="https://komarev.com/ghpvc/?username=EmTeeVee" width="1" height="1" />
