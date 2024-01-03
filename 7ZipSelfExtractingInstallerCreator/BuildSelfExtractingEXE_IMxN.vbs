Option Explicit

Const SCRIPT_VERSION = "1.3"

' variables
Dim fso
Dim logFile
Dim solutionFolder
Dim buildFolder
Dim releaseFolder
Dim baseReleaseFolder
Dim useSolutionNameInReleaseFolder
Dim tempDirFileName
Dim tempDirFile
Dim fileListFileName
Dim fileListFile
Dim fileEntry
dim fileObj
Dim exeName
Dim useVersion
Dim useTrial
Dim exeFullPath
Dim releaseVersion
Dim batchFileName
Dim batchFile
Dim tempArchiveName
Dim selfExtractorName
Dim sfxFileName
dim configFileName
dim configFile
Dim shell
Dim shellCmd
Dim tmp1
Dim tmp2
dim i
dim localFiletime
dim utcFiletime
dim datetime

Set datetime = CreateObject("WbemScripting.SWbemDateTime")

' default build parameters
useVersion = False
useTrial = False
baseReleaseFolder = "C:\Program Files\JDSU\"
useSolutionNameInReleaseFolder = True

' create the FileSystemObject for file and folder handling
Set fso = CreateObject("Scripting.FileSystemObject")

' shell object for command-line execution
Set shell = WScript.CreateObject("WScript.shell")

' get the solution folder from the arguments passed to the script file
' this is where the 7-Zip tool is located
solutionFolder = WScript.Arguments.Item(0)

' create the log file
Set logFile = fso.CreateTextFile(solutionFolder & "BuildSelfExtractingEXE.log")
logFile.WriteLine "SCRIPT_VERSION=" & SCRIPT_VERSION
logFile.WriteLine "============================================================================================"
logFile.WriteLine "solutionFolder=" & solutionFolder

' get the build folder from the arguments passed to the script file
' this is where we get the application files from
buildFolder = WScript.Arguments.Item(1)
logFile.WriteLine "buildFolder=" & buildFolder

' get the EXE name from the arguments passed to the script file
exeName = WScript.Arguments.Item(2)
logFile.WriteLine "exeName=" & exeName

' parse the remaining arguments
logFile.WriteLine "Found a total of " & WScript.Arguments.Count & " command-line arguments"
For i = 3 to WScript.Arguments.Count-1
	tmp1 = WScript.Arguments.Item(i)
	logFile.WriteLine "argIndex " & i & ": " & tmp1
	tmp2 = Split(tmp1,"=")
	If UCase(tmp2(0)) = "USEVERSION" Then
		If UCase(tmp2(1)) = "TRUE" Then
			useVersion = True
		Else
			useVersion = False
		End If
	ElseIf UCase(tmp2(0)) = "RELEASEFOLDER" Then
		useSolutionNameInReleaseFolder = False
		baseReleaseFolder = tmp2(1)
		If Right(baseReleaseFolder, 1) <> "\" Then
			baseReleaseFolder = baseReleaseFolder & "\"
		End If
	ElseIf UCase(tmp2(0)) = "TRIAL" Then
		If UCase(tmp2(1)) = "TRUE" Then
			useTrial = True
		Else
			useTrial = False
		End If
	Else
		logFile.WriteLine "WARNING: Unrecognized command line parameter"
	End If
Next

' get the version of the EXE
exeFullPath = buildFolder & exeName
logFile.WriteLine "exeFullPath= " & exeFullPath
releaseVersion = fso.GetFileVersion(exeFullPath)
logFile.WriteLine "releaseVersion=" & releaseVersion
' truncate the release version to only major for use in folder/file naming
tmp1 = Split(releaseVersion,".")
releaseVersion = tmp1(0)

' create the release folder name in Program Files\JDSU based on the EXE name and version
' this is where the application will be installed when extracted
If useVersion Then
	If useSolutionNameInReleaseFolder Then
		releaseFolder = baseReleaseFolder & Left(exeName, Len(exeName)-4) & "_" & releaseVersion
	Else
		releaseFolder = left(baseReleaseFolder,len(baseReleaseFolder)-1)  ' strip the trailing \
	End If
Else
	If useSolutionNameInReleaseFolder Then
		releaseFolder = baseReleaseFolder & Left(exeName, Len(exeName)-4)
	Else
		releaseFolder = left(baseReleaseFolder,len(baseReleaseFolder)-1)  ' strip the trailing \
	End If
End If

If useTrial Then
	releaseFolder = releaseFolder & "_Trial"
End If

logFile.WriteLine "releaseFolder=" & releaseFolder

' delete the release folder if it exists
If fso.FolderExists(releaseFolder) Then
	logFile.WriteLine "Deleting existing release folder " & releaseFolder
	fso.DeleteFolder releaseFolder, True
End If

' copy all from the build folder to the release folder
' do this before creating the self-extractor because we want the self-extract to extract to this folder, not the build folder
logFile.WriteLine "Copying build output to release folder"
tmp1 = Split(releaseFolder, "\")
tmp2 = tmp1(0)
For i = 1 to UBound(tmp1)
	If Len(tmp1(i)) > 0 Then
		tmp2 = tmp2 & "\" & tmp1(i)
		If Not fso.FolderExists(tmp2) Then
			fso.CreateFolder tmp2
		End If
	End If
Next
fso.CopyFolder Left(buildFolder, Len(buildFolder)-1), releaseFolder   ' remove trailing \

' create a file containing a list of all files in the release folder
logFile.WriteLine "Getting list of files in the release folder"
tempDirFileName = solutionFolder & "tempdir.txt"
shellCmd = "cmd.exe /C dir " & Chr(34) & releaseFolder & Chr(34) & " /s /b /a-d > " & Chr(34) & tempDirFileName & Chr(34)
logFile.WriteLine "Running shell command '" & shellCmd & "'"
shell.Run shellCmd, 7, True
' add size and timestamp (in UTC time) to list of files
fileListFileName = releaseFolder & "\FileList.csv"
Set fileListFile = fso.CreateTextFile(fileListFileName)
Set tempDirFile = fso.OpenTextFile(tempDirFileName)
fileListFile.WriteLine "FILE,SIZE,DATE"
While Not tempDirFile.AtEndOfStream
	fileEntry = tempDirFile.ReadLine
	If Len(fileEntry) > 0 Then
		Set fileObj = fso.GetFile(Trim(fileEntry))
		localFiletime = fileObj.DateLastModified
		datetime.SetVarDate(localFiletime)
		utcFiletime = datetime.GetVarDate(false)
		fileEntry = Replace(FileEntry, releaseFolder & "\", "") 'make relative path
		LogFile.WriteLine fileEntry & "," & localFiletime & "," & utcFiletime
		fileEntry = fileEntry & "," & fileObj.Size & "," & utcFileTime
		fileListFile.WriteLine fileEntry
	End If
Wend
tempDirFile.Close
fileListFile.Close

' create the batch file that will auto-launch the executable after extraction
' needed because the self-extractor is running from the parent folder
' the batch file corrects this
batchFileName = releaseFolder & "\launcher.bat"
Set batchFile = fso.CreateTextFile(batchFileName)
batchFile.WriteLine "cd " & releaseFolder
batchFile.WriteLine "start " & exeName
batchFile.WriteLine "exit"
batchFile.Close

' create the zip file containing the contents of the release folder
tempArchiveName = solutionFolder & "temp.7z"
logFile.WriteLine "Creating the temporary archive '" & tempArchiveName & "'"
If fso.FileExists(tempArchiveName) Then
	fso.DeleteFile(tempArchiveName)
End If
shellCmd = Chr(34) & solutionFolder & "7ZipSelfExtractingInstallerCreator\7za.exe" & Chr(34) & " a " & Chr(34) & tempArchiveName & Chr(34) & " " & Chr(34) & releaseFolder & Chr(34)
logFile.WriteLine "Running shell command '" & shellCmd & "'"
shell.Run shellCmd, 7, True

' create the config file for the self-extractor
logFile.WriteLine "Creating config file for self-extractor"
configFileName = solutionFolder & "7ZipSelfExtractingInstallerCreator\7z_config.txt"
Set configFile = fso.CreateTextFile(configFileName)
configFile.WriteLine ";!@Install@!UTF-8!"
configFile.WriteLine "InstallPath=" & Chr(34) & Replace(baseReleaseFolder,"\","\\") & Chr(34)
configFile.WriteLine "RunProgram=" & Chr(34) & " " & Chr(34)
' -- to start the deployed application automatically, comment out the above line and uncomment the line below --
'configFile.WriteLine "RunProgram=" & Chr(34) & "nowait:" & "\" & Chr(34) & Replace(batchFileName,"\","\\") & "\" & Chr(34) & Chr(34)
configFile.WriteLine "GUIMode=" & Chr(34) & "2" & Chr(34)
configFile.WriteLine ";!@InstallEnd@!"
configFile.Close

' build the self-extractor
sfxFileName = solutionFolder & "7ZipSelfExtractingInstallerCreator\7zsd.sfx"
selfExtractorName = ""
If useVersion Then
	selfExtractorName = "_" & releaseVersion 
End If
If useTrial Then
	selfExtractorName = selfExtractorName & "_Trial"
End If
selfExtractorName = solutionFolder & Left(exeName, Len(exeName)-4) & selfExtractorName & "_installer.exe"
If fso.FileExists(selfExtractorName) Then
	fso.DeleteFile(selfExtractorName)
End If
logFile.WriteLine "Creating the self-extracting executable '" & selfExtractorName & "'"
shellCmd = "cmd.exe /C copy /b " & Chr(34) & sfxFileName & Chr(34) & " + " & Chr(34) & configFileName & Chr(34) & " + " & Chr(34) & tempArchiveName & Chr(34) & " " & Chr(34) & selfExtractorName & Chr(34) 
logFile.WriteLine "Running shell command '" & shellCmd & "'"
shell.Run shellCmd, 7, True

' clean up
fso.DeleteFile tempDirFileName, True
'fso.DeleteFile configFileName, True
fso.DeleteFile tempArchiveName, True

logFile.WriteLine "Done."

