
saveRoot = "C:\Users\Administrator\AppData\Roaming\EldenRing\76561197960267366\"

LOAD_TO_BACKUP = FALSE
RESEARCH = FALSE

MAX_SAVES = 4

Function Speak(text)
  set speaker = CreateObject("SAPI.SpVoice")
  speaker.Volume = 37
  speaker.speak text
End Function

Function SaveFileMain(index)
  SaveFileMain = saveRoot + "save" + cstr(index) + ".sav"
End Function

Function SaveFileBackup(index)
  SaveFileMain = saveRoot + "save" + cstr(index) + ".bak"
End Function

Function ReadAge(age)
  age = -age
  if age > 60 then
	Speak("Age is over " & (age \ 60) & "minutes")
  else
	Speak("Age is " & age & "seconds")
  end if
End Function

task = ""
Set oArgs = WScript.Arguments
    For Each s In oArgs
		task = s
    Next
Set oArgs = Nothing

currentTime = Now


saveDoc = saveRoot + "ER0000.sl2"
saveDocBk = saveRoot + "ER0000.sl2.bak"

Set fso = CreateObject("Scripting.FileSystemObject")

Set mainfile = fso.GetFile(saveDoc)
saveTime = mainfile.DateLastModified

Set backupfile = nothing

IF RESEARCH THEN
	Set backupfile = fso.GetFile(saveDocBk)
	backupTime = backupfile.DateLastModified

	age = DateDiff("s", saveTime, backupTime)

	'Speak("Backup age is " & age) ' Should be zero
end if

recentIndex = -1
recentTime = -1

alias = Array("Recent", "Old", "Older", "Oldest")

FOR i = 0 to MAX_SAVES - 1
	'Speak("Checking " & i)
    opFile = SaveFileMain(i)
    IF fso.FileExists(opFile) THEN
		'Speak("Exists " & i)
		Set file = fso.GetFile(opFile)
		modTime = file.DateLastModified	
		
		set file = nothing
		
		IF recentTime < 0 THEN
			recentTime = modTime
			recentIndex = i
		END IF 
		
		if modTime > recentTime THEN
			recentTime = modTime
			recentIndex = i
		END IF
	END IF

Next

'recentIndex = -1

writeIndex = (recentIndex + 1) mod MAX_SAVES
'Speak("Write Index is " & writeIndex)

if recentIndex < 0 then
	recentIndex = 0
end if

'Speak("Recent Index is " & recentIndex)

'Speak(task)
'Speak(Time)


if 0 = StrComp(task, "save") then
	age = DateDiff("s", saveTime, recentTime)
	if age < 0 Then
		age = DateDiff("s", currentTime, saveTime)
		ReadAge(age)
		mainfile.copy(SaveFileMain(writeIndex))
		if err.number=0 then
			Speak("Saved to " & writeIndex)
		else
			Speak("Error saving. Please try again")
		end if
	ELSE
		Speak("Not Saved")
	END if
end if

if 0 = StrComp(task, "load-1") then
	task = "load"
	recentIndex = (recentIndex - 1 + MAX_SAVES) mod MAX_SAVES
end if

if 0 = StrComp(task, "load-2") then
	task = "load"
	recentIndex = (recentIndex - 2 + MAX_SAVES) mod MAX_SAVES
end if

if 0 = StrComp(task, "load-3") then
	task = "load"
	recentIndex = (recentIndex - 3 + MAX_SAVES) mod MAX_SAVES
end if

if 0 = StrComp(task, "load") then
	opfile = SaveFileMain(recentIndex)
	if fso.FileExists(opFile) then
		set source = fso.GetFile(opfile)
		source.copy(saveDoc)
		mainOK = (err.number=0)

		if LOAD_TO_BACKUP and RESEARCH then
			source.copy(saveDocBk)
			backupOK = (err.number=0)
		ELSE
			backupOK = TRUE
		END if

		if mainOK and backupOK then
			Speak("Loaded from " & recentIndex)
		else
			Speak("Error loading. Please try again")
		end if
		
		set source = nothing
	else
		Speak("Saves Not Found")
	end if
end if

'Speak(saveTime)
Set file = Nothing
Set fso = Nothing
set mainfile = nothing

