# randomlyPlayFolder
VBScript to generate a playlist from a folder with random order and then run the player on the generated playlist.

3 Usages : 
- run `wscript randomlyPlayFolder.vbs <recursive:true|false> <folder path> <extensions separate by a comma> <player path> `.
	ex : *wscript randomlyPlayFolder.vbs true D:\mp3 mp3,MP3,m4a "C:\Program Files (x86)\Winamp\winamp.exe"*
- edit the randomlyPlayFolder.vbs to set the folder, the recursivity, file extensions to load and the player path to run.
- copy and edit the samplePlay
