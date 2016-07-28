Dim FSO, OWS, oWMI, oLog, cFolders, oFolder
Dim sRoot, sUserProfile, sLogFile, sGroup, sDomain

Set FSO = CreateObject("Scripting.FileSystemObject")
Set OWS = CreateObject("Wscript.Shell")
Set oWMI = GetObject("winmgmts:\\.\root\cimv2") 

sRoot = fGetFolder
Set oFolder = FSO.GetFolder(sRoot)
sUserProfile = OWS.ExpandEnvironmentStrings("%USERPROFILE%")
sLogFile = sUserProfile & "\Desktop\413FLTS_Permissions.csv"

sGroup = "HA-413-FLTS-Lead-CSA"
sDomain = "HURLBURT"

'Const VBTEXTCOMPARE = 1
'Const FORAPPENDING = 8

If FSO.FileExists(sLogFile) Then
	Set oLog = FSO.OpenTextFile(sLogFile, 8)
	oLog.WriteBlankLines(1)
	oLog.WriteLine(String(150, "-"))
	oLog.WriteBlankLines(1)
Else
	Set oLog = FSO.CreateTextFile(sLogFile)
End If

oLog.WriteLine("Fix Permissions on Folders")
oLog.WriteLine("Root Folder: " & sRoot)
oLog.WriteLine("Started: " & Now)
oLog.WriteBlankLines(1)
oLog.WriteLine(",Status,Error Number,Folder")

On Error Resume Next

EvaluatePermissions oFolder

oLog.WriteBlankLines(1)
oLog.WriteLine("Ended: " & Now)

'--------------------------------------------------------------------
' Functions and Subroutines
'--------------------------------------------------------------------
	'--------------------------------------------------------------------
	' Function to return the folder path to be evaluated.
	'--------------------------------------------------------------------
	Function fGetFolder()
		Dim sFolder
		
		sFolder = InputBox("Please input the folder path to be evaluated." &_
			VbCr & VbCr & "Use the format: DriveLetter:\Folder\Subfolder" &_
			VbCr & "Example: C:\Windows\System32" & VbCr & VbCr & "To " &_
			"cancel, leave blank or hit the ""Cancel"" button.", "Folder")
			
		If sFolder = "" Then Wscript.Quit
		If Not FSO.FolderExists(sFolder) Then
			Wscript.Echo "The folder path you enterred was invalid."
			fGetFolder
		Else
			fGetFolder = sFolder
		End If
	End Function
	
	'--------------------------------------------------------------------
	' Subroutine to evaluate permissions on the target folder.
	'--------------------------------------------------------------------
	Sub EvaluatePermissions(Folder)
		Dim oFldr, oSub, iFlag, oSD, oACE
		
		For Each oSub in Folder.Subfolders
			iFlag = 0
			Set oFldr = oWMI.Get("Win32_LogicalFileSecuritySetting='" & oSub.Path & "'")
			If oFldr.GetSecurityDescriptor(oSD) = 0 Then
				For Each oACE in oSD.DACL
					If UCase(oACE.Trustee.Name) = UCase(sGroup) Then
						iFlag = iFlag + 1
					End If
				Next
				Set oACE = Nothing
			End If
			Set oSD = Nothing
			Set oFldr = Nothing
			
			If iFlag = 0 Then
				FixPermissions oSub
			End If
			
			EvaluatePermissions oSub
		Next
	End Sub
	
	'--------------------------------------------------------------------
	' Subroutine to fix permissions on the target folder
	'--------------------------------------------------------------------
	Sub FixPermissions(Folder)
		Dim iRunErr, oFile, oSub
		
		iRunErr = OWS.Run("cmd.exe /c cacls """ & Folder.Path & """ /e /c /g " &_
			sDomain & "\" & sGroup & ":C", 2, True)
		WriteLog iRunErr, Folder.Path
		iRunErr = 0
			
		For Each oFile in Folder.Files
			iRunErr = OWS.Run("cmd.exe /c cacls """ & oFile.Path & """ /e /c /g " &_
				sDomain & "\" & sGroup & ":C", 2, True)
			WriteLog iRunErr, oFile.Path
			iRunErr = 0
		Next
		
		For Each oSub in Folder.SubFolders
			FixPermissions oSub
		Next
	End Sub
	
	'--------------------------------------------------------------------
	' Function to write a line to the log file
	'--------------------------------------------------------------------
	Sub WriteLog(ErrorNumber, FolderPath)
		
		If ErrorNumber = 0 Then
			oLog.WriteLine(",FIXED,0," & FolderPath)
		Else
			oLog.WriteLine(",ERROR," & ErrorNumber & "," & FolderPath)
		End If
		
	End Sub