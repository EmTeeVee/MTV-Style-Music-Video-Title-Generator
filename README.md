# MTV-Style-Music-Video-Title-Generator
Creates MTV Style Title Cards for music videos using their filenames
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
also, the MTV style font "Kabel-Black.ttf" should be available, ideally in the folder you are muxing in (the script will prompt you for it, if it is not there), as it is added to the muxed file  

I built this so I can play music videos with title cards in Plex (the release year after running this script should be visible in Plex, so sorting by year or decade is not a problem anymore), but please be aware that because of the way Plex works, it usually transcodes the ass-subtitles contained in the final *.mkv files. If you use an NVidia shield or the Plex Desktop App it should work fine, in other scenarios the transcoding might cause problems.

Music videos must be named like this:  
"artist - title (remix) [original artist] (year, album) [extras].videoextension"  
(" - " has to be between artist and title, "(year)" or "(year, ...)" is the second important distinguisher), everything else is optional  
at least "artist - title.ext" has to be present in every single file in the chosen directory (and its subdirectories), otherwise the program will not run

Examples:  
Siouxsie And The Banshees - Kiss Them For Me (1991, Superstition).vob  
ABBA - I Can Be That Woman (2021, Voyage) [Lyric Video].mp4  
Blondie - The Tide Is High [The Paragons] (1980, Autoamerican).mkv  

Additionally, the script can add subtitle tracks to the music videos if you have them present alongside your music video file and named accordingly, for example:  
Eurythmics - Here Comes The Rain Again (1983, Touch).mp4  
Eurythmics - Here Comes The Rain Again (1983, Touch)_eng.srt  
Eurythmics - Here Comes The Rain Again (1983, Touch)_fre.srt  


Finally, some information is written at the end of the filename in wavy brackets {} after muxing, so that {SD}, {4K}, {AV1} or {VP9} files are marked. This is for quality and compatibilty reasons that one might have to consider.
<img src="https://komarev.com/ghpvc/?username=EmTeeVee" width="1" height="1" />
