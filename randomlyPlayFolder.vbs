rem On Error Resume Next

Dim playerPath
Dim tempPlayList
Dim sleepTime
Dim startingFolder
Dim authorizedExtension
Dim recursive
Dim deleteM3U
REM CONF
deleteM3U=true
recursive=false
startingFolder="D:\Mp3"
authorizedExtension="mp3, MP3"
sleepTime=30000
playerPath="C:\Program Files (x86)\Winamp\winamp.exe"
REM playerPath="C:\Program Files (x86)\iTunes\iTunes.exe"

Set WshShellObj = WScript.CreateObject("WScript.Shell")
Set WshProcessEnv = WshShellObj.Environment("Process") 

tempPlayList=WshProcessEnv("TEMP")&"\list.m3u"
REM from argument
If wscript.Arguments.length > 0 Then 
  recursive=(wscript.arguments(0)="true")
End if
If wscript.Arguments.length > 1 Then 
  startingFolder=wscript.arguments(1)
End if
If wscript.Arguments.length > 2 Then 
  authorizedExtension=wscript.arguments(2)
End if
If wscript.Arguments.length > 3 Then 
  playerPath=wscript.arguments(3)
End if
If wscript.Arguments.length > 4 Then 
  tempPlayList=wscript.arguments(4)
  deleteM3U=false
End if


Dim tab()
Dim size
Dim cpt
cpt=0
size=0

Sub count(ByVal sFolder)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
set f=fso.getfolder(sFolder)
for each fichier in f.Files
  extension=fso.GetExtensionName(fichier.path)
  if Len(extension)>0 and InStrRev(authorizedExtension, extension) > 0 then 
      size = size+1
  end if
next
if (recursive) then
	set sf=f.subfolders
	for each f1 in sf
		count(f1)
	next
end if
End Sub

count(startingFolder)

ReDim tab(size)

Sub browse(ByVal sFolder)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
set f=fso.getfolder(sFolder)
for each fichier in f.Files
	extension=fso.GetExtensionName(fichier.path)
     if Len(extension)>0 and InStrRev(authorizedExtension,extension) > 0 then 
      tab(cpt)=fichier.path
      cpt=cpt+1
       end if
next
if (recursive) then
	set sf=f.subfolders
	for each f1 in sf
		browse f1
	next
end if
End Sub

browse startingFolder

rem msgbox size
Set fsod = CreateObject("Scripting.FileSystemObject")

Dim ftemp

Set ftemp = fsod.OpenTextFile(tempPlayList, 2,true)

Dim random
Dim ret2
ret2=""
For i = 1 To size
cpt=0
cpt2=0
placed=False
Randomize
random=Int( ( size-1 ) * Rnd )
Dim placed
placed=false
Dim fsod
Dim fde
Dim cpt3
cpt3=0

On Error Resume Next
do while not placed
	rem	msgbox "random : "&random&"/cpt :" & cpt & "/cpt2 :" & cpt2 & "/"&tab(cpt)& vbnewline
   if (tab(cpt) <> "" ) then 
      if (cpt2 = random) then
			rem msgbox tab(cpt)
			ftemp.writeline(tab(cpt))
			If Err.Number <> 0 Then
				WScript.Echo "Le fichier suivant n'a pas pu être ajouté : "& tab(cpt) & " : "& Err.Number & Err.Description
				Err.Clear
			End If
            tab(cpt)=""
            size = size - 1
            placed=True
      end if
	  cpt2=cpt2+1
	else
	   cpt3=cpt3+1
	   rem msgbox i&" "&cpt3 & "cpt"&cpt
     end if
	cpt = cpt +1
 loop
Next 


ftemp.close()

if (cpt=0) then 
  msgbox "Aucun fichier"
else
  set WshShell = createObject("WScript.shell")
  rem msgbox """"&playerPath&"""" & " " & tempPlayList
  if deleteM3U then 
	Wshshell.run """"&playerPath&"""" & " " & tempPlayList
	end if
  Wscript.Sleep sleepTime
  if deleteM3U and fsod.FileExists (tempPlayList) then
    fsod.DeleteFile(tempPlayList) 
  end if
end if