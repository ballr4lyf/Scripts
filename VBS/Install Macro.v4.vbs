Option Explicit

'Strings and Integers
Dim iAnswer, sUserProfile, sDest, sSource

'Objects
Dim OWS, FSO, oWMIService, oProcess, oFolder, objFile
Dim oRegExp, NewFile

'Collections
Dim cProcesses, cFiles

'On Error Resume Next

Set OWS = CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

'--------------------------------------------------------------------------
'This section will open a Message box and give the user an option to 
'continue with deployment or to cancel.
'--------------------------------------------------------------------------
iAnswer = _
	MsgBox("Outlook will be closed if it is currently running" & VbCrLf &_
		"Please save all your work before continuing." & VbCrLf & VbCrLf & "Outlook" &_
		" will be opened after the install/repair is complete." & VbCrLf & VbCrLf &_
		"Do you wish to continue?", vbYesNo, "IT Help Macro Install/Repair")

'--------------------------------------------------------------------------
'Cancel script if user clicks on "No"
'--------------------------------------------------------------------------
If Not iAnswer= vbYes Then
	OWS.Popup "Install/Repair Cancelled.", 3, "Cancelled"
	Wscript.Quit
End If

'--------------------------------------------------------------------------
'This section will terminate the Outlook process.
'--------------------------------------------------------------------------
Set oWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set cProcesses=oWMIService.ExecQuery _
	("SELECT * FROM Win32_Process WHERE Name = 'OUTLOOK.EXE'")
For Each oProcess in cProcesses
	oProcess.Terminate()
Next

Wscript.Sleep 5000

'--------------------------------------------------------------------------
'This section copies VbaProject.OTM from the current folder to the user's
'profile.
'--------------------------------------------------------------------------
sUserProfile = OWS.ExpandEnvironmentStrings("%USERPROFILE%")
sDest = sUserProfile & "\AppData\Roaming\Microsoft\Outlook"
sSource = OWS.CurrentDirectory & "\Macro"

'--------------------------------------------------------------------------
'Will check for and then delete any files named "VbaProject.old"
'--------------------------------------------------------------------------
If FSO.FileExists (sDest & "\VbaProject.old") Then
	FSO.DeleteFile (sDest & "\vbaProject.old")
End if

'--------------------------------------------------------------------------
'Will change the file extension for any ".OTM" files to ".old".  This will
'allow us to recover the user's macros in case of a problem.
'--------------------------------------------------------------------------
Set oFolder = FSO.GetFolder(sDest)
Set cFiles = oFolder.Files
	Set oRegExp = New RegExp
	oRegExp.Pattern = ".OTM"
	oRegExp.IgnoreCase = True
For Each objFile in cFiles
	if oRegExp.Test(objFile.Name) Then
		NewFile = oRegExp.Replace (objFile.Name, ".old")
		FSO.MoveFile objFile, sDest & "\" & NewFile
	End if
Next

'--------------------------------------------------------------------------
'Copies the "VbaProject.OTM" file from the source directory
'--------------------------------------------------------------------------
If FSO.FileExists (sSource & "\VbaProject.OTM") Then
	FSO.CopyFile (sSource & "\VbaProject.OTM"),(sDest & "\")
End if

'--------------------------------------------------------------------------
'Opens Outlook
'--------------------------------------------------------------------------
OWS.Run "Outlook"

'--------------------------------------------------------------------------
'This sections informs the user that the install/repair has completed.  
'The popup remains on screen for 10 seconds.
'--------------------------------------------------------------------------
CreateObject("Wscript.Shell").Popup "Outlook Macro setup is complete.", _
	10, "Complete"

'--------------------------------------------------------------------------
'Cleanup
'--------------------------------------------------------------------------
Set oRegExp = Nothing
Set cFiles = Nothing
Set oFolder = Nothing
Set cProcesses = Nothing
Set FSO = Nothing
Set OWS = Nothing