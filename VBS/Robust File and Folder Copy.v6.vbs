'==========================================================================
'
'  Use "Robocopy" with customizeable Source and Destination paths, as well 
'     as standardized options, including multithreaded copy.
'
'                     **** REQUIRES WINDOWS 7 ****
'
' Created by: Robert D. Rathbun
' Date:  11/18/2011
' Modified Date:  2/2/2012
'
'==========================================================================

'--------------------------------------------------------------------
'Requires that variables be declared using the "Dim" statement before
' they are used.
'--------------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------
Dim sSource, sDest, sLogFile, sWinDir, sRobocopy, sOptions, iAnswer
Dim OWS, FSO, oRegExp, sUsrProfile, oLogFile

Set OWS = CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

sUsrProfile = OWS.ExpandEnvironmentStrings("%USERPROFILE%")

'--------------------------------------------------------------------
'Call the "VerifySource" subroutine
'--------------------------------------------------------------------
VerifySource

'--------------------------------------------------------------------
'Call the "VerifyDestination" subroutine
'--------------------------------------------------------------------
VerifyDestination

'--------------------------------------------------------------------
'Create the "robocopy.log" log file in the destination folder.
'--------------------------------------------------------------------
sLogFile = sUsrProfile & "\desktop\robocopy_" & Year(Date) & "-" & Month(Date) & Day(Date) & ".log"

sWinDir = OWS.ExpandEnvironmentStrings("%WINDIR%")
sRobocopy = sWinDir & "\System32\Robocopy.exe "

'--------------------------------------------------------------------------
' Testing for spaces in the paths using Regular Expressions.
'   (Paths need to be enclosed in quotation marks if spaces are present)
'--------------------------------------------------------------------------
Set oRegExp = New RegExp
	With oRegExp
		.Pattern = " "
		.IgnoreCase = True
	End With

  If oRegExp.Test(sLogFile) Then
    sLogFile = Chr(34) & sLogFile & Chr(34)
  End If
  
  If oRegExp.Test(sSource) Then
    sSource = Chr(34) & sSource & Chr(34)
  End If
  
  If oRegExp.Test(sDest) Then
    sDest = Chr(34) & sDest & Chr(34)
  End If

'--------------------------------------------------------------------------
'Determine if the user wants to use the "Purge" option of robocopy and set
' the standard options
'--------------------------------------------------------------------------
iAnswer = MsgBox("Do you want to purge the data from the Destination directory?" &_
    VbCrLf & "(Data will be deleted from the destination directory if it no " &_
    "longer exists at the source)", vbYesNo, "Purge")
  If iAnswer = vbYes Then
    sOptions = "/E /PURGE /MT:50 /R:1 /W:1 /V /NP /LOG+:" & sLogFile
  Else
    sOptions = "/E /MT:50 /R:1 /W:1 /V /NP /LOG+:" & sLogFile
  End If

'--------------------------------------------------------------------------
'Run the robocopy command and move the logfile
'--------------------------------------------------------------------------
Dim oDestLogFile

If FSO.FileExists(sDest & "\robocopy.log") Then
	Set oDestLogFile = FSO.GetFile(sDest & "\robocopy.log")
	If oDestLogFile.Size > 300000000 Then
		FSO.MoveFile (sDest & "\robocopy.log"), sUsrProfile & "\robocopy.old.log"
		Set oDestLogFile = FSO.GetFile(sUsrProfile & "\robocopy.old.log")
	Else
		FSO.MoveFile sDest & "\robocopy.log", sUsrProfile & "\"
	End If
End If

OWS.Run("cmd.exe /c " & sRobocopy & sSource & " " & sDest & " " & sOptions)

Set oLogFile = FSO.GetFile(sUsrProfile & "\robocopy.log")

If FSO.FileExists(sDest & "\robocopy.log") OR _
	FSO.FileExists(sDest & "\robocopy.old.log") Then
		FSO.MoveFile oDestLogFile.Path, sUsrProfile & "\"
End If

If FSO.FileExists(oLogFile.Path) Then
	oLogFile.Move sDest & "\"
End If

	'--------------------------------------------------------------------------
	' Subroutine to verify that the source path is a valid folder path.
	'--------------------------------------------------------------------------
	Sub VerifySource
	sSource = InputBox("Please input the source folder path." & VbCr & VbCr &_
		"If left blank, the script will exit.", "Source")
	    If Not FSO.FolderExists(sSource) Then
			If sSource = "" then
				Wscript.Quit
			End If
			Wscript.Echo "The folder provided (" & sSource & ") does not exist."
			VerifySource
	    End If
	End Sub

	'--------------------------------------------------------------------------
	' Subroutine to verify that the destination path is a valid folder path.
	'--------------------------------------------------------------------------
	Sub VerifyDestination
	sDest = InputBox("Please input the destination folder path."& VbCr & VbCr &_
		"If left blank, the script will exit.", "Destination")
		If Not FSO.FolderExists(sDest) Then
			If sDest = "" Then
				Wscript.Quit
			End If
			Wscript.Echo "The folder provided (" & sDest & ") does not exist."
			VerifyDestination
		End If
	End Sub

'--------------------------------------------------------------------------
'Cleanup
'--------------------------------------------------------------------------
Set oDestLogFile = Nothing
Set oRegExp = Nothing
Set oLogFile = Nothing
Set FSO = Nothing
Set OWS = Nothing