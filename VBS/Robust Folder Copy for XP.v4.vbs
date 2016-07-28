'--------------------------------------------------------------------
' Script to perform robust folder copies on systems without Robocopy
'   available (i.e. Windows XP).
'
' Created By:  Robert Rathbun
' Created Date:  12/15/2011
' Modified Date:  2/2/2012
'
'--------------------------------------------------------------------

'--------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------
Option Explicit

'Objects
Dim FSO, OWS, oSource, oDest, oLog, oSrcfile, oDstFile, oSrcSub, oDstSub, oLogRen, oXl

'Strings
Dim sSource, sDest, sInput, sFolderType, sTemp, sLogName, sLogFile, sTempMove, sPurge
Dim sDstFileSize, sDstFilePath, sDstSubSize, sDstSubPath, Size, Suffix
Dim TargetLen, RemoveSection, TrimFrom, RemSection, TrimLen

'Integers
Dim iPurge, iPurgeFldr, iPurgeFile, iReplaceFldr, iReplaceFile, iCopyFldr, iCopyFile
Dim iSkipFldr, iSkipFile, iErrFldr, iErrFile

'--------------------------------------------------------------------
' Script Start
'--------------------------------------------------------------------

Set FSO = CreateObject("Scripting.FileSystemObject")
Set OWS = Wscript.CreateObject("Wscript.Shell")

'--------------------------------------------------------------------
' Instruct user to provide folders.
'--------------------------------------------------------------------
sSource = fVerifyFolders("Source")
	Set oSource = FSO.GetFolder(sSource)
sDest = fVerifyFolders("Destination")
	Set oDest = FSO.GetFolder(sDest)

'--------------------------------------------------------------------
' Determine if user wants to purge data from destination.
'--------------------------------------------------------------------
iPurge = MsgBox("Do you want to ""Purge"" data from the Destination folder?" _
	& VbCr & VbCr & "WARNING: Data will be deleted from the Destination folder " _
	& "if it does not exist at the source.", VbYesNo, "Purge")

'--------------------------------------------------------------------
' Create log file and write header lines to log file
'--------------------------------------------------------------------
sTemp = OWS.ExpandEnvironmentStrings("%TEMP%")
sLogName = "CopyLog.csv"
sLogFile = sTemp & "\" & sLogName
	If Not FSO.FileExists(sLogFile) Then
		Set oLog = FSO.CreateTextFile(sLogFile)
	Else
		Set oLog = FSO.GetFile(sLogFile)
		oLog.Delete
		Set oLog = FSO.CreateTextFile(sLogFile, 8)
	End If
If FSO.FileExists(sDest & "\" & sLogName) Then
	sTempMove = sTemp & "\" & Year(Now) & "-" & Month(Now) & "-" &_
		Day(Now) & "_" & sLogName
	FSO.MoveFile sDest & "\" & sLogName, sTempMove
End If

oLog.WriteBlankLines(1)
oLog.WriteLine(String(150, "-"))
oLog.WriteLine("Folder Copy Utility for XP")
oLog.WriteLine("Started: " & Now)
oLog.WriteLine("Source Folder: " & sSource)
oLog.WriteLine("Destination Folder: " & sDest)
	If iPurge = VbYes Then
		sPurge = "Yes"
	ElseIf iPurge = VbNo Then
		sPurge = "No"
	End If
oLog.WriteLine("Purge Enabled: " & sPurge)
oLog.WriteLine("NOTE: Purge will remove data from the destination folder if it does not")
oLog.WriteLine("       exist at the source.")
oLog.WriteLine("Log Location: " & sDest & "\" & sLogName)
oLog.WriteBlankLines(1)
oLog.WriteLine(",Action,Error,Size,Error No.,Path")
	'Actions = Purged, Replaced, Copied, Skipped, Error

'--------------------------------------------------------------------
' Counting Actions Performed
'--------------------------------------------------------------------
iPurgeFldr = 0
iPurgeFile = 0
iReplaceFldr = 0
iReplaceFile = 0
iCopyFldr = 0
iCopyFile = 0
iSkipFldr = 0
iSkipFile = 0
iErrFldr = 0
iErrFile = 0

'--------------------------------------------------------------------
' Main Process
'--------------------------------------------------------------------
'On Error Resume Next

RecursiveCopy oSource, oDest

'On Error GoTo 0

'--------------------------------------------------------------------
' Write Summary lines of log file and copy file to destination folder
'--------------------------------------------------------------------
oLog.WriteBlankLines(1)
oLog.WriteLine(",,Purged,Copied,Replaced,Skipped,Errors")
oLog.WriteLine(",Folders," & iPurgeFldr & "," & iCopyFldr & "," & iReplaceFldr & "," & iSkipFldr &_
	"," & iErrFldr)
oLog.WriteLine(",Files," & iPurgeFile & "," & iCopyFile & "," & iReplaceFile & "," & iSkipFile &_
	"," & iErrFile)
oLog.WriteLine(",Totals," & iPurgeFldr + iPurgeFile & "," & iCopyFldr + iCopyFile & "," &_
	iReplaceFldr + iReplaceFile & "," & iSkipFldr + iSkipFile & "," & iErrFldr + iErrFile)
oLog.WriteBlankLines(1)
oLog.WriteLine("Copy Completed: " & Now)
oLog.Close

If Not FSO.FileExists(sDest & "\" & sLogName) Then
	FSO.CopyFile sLogFile, sDest & "\"
Else
	Set oLogRen = FSO.GetFile(sDest & "\" & sLogName)
	oLogRen.Name = Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "_" & sLogName
	FSO.CopyFile sLogFile, sDest & "\"
End If

If Not sTempMove = VbNullString Then
	FSO.MoveFile sTempMove, sDest & "\"
End If

FSO.DeleteFile(sLogFile)

Set oXl = CreateObject("Excel.Application")
  oXl.Application.Workbooks.Open sDest & "\" & sLogName
  oXl.Application.Visible = True

'--------------------------------------------------------------------
' Subroutines and Functions
'--------------------------------------------------------------------
	'--------------------------------------------------------------------
	' fVerifyFolders Function (Verify that folders exist)
	'--------------------------------------------------------------------
	Function fVerifyFolders(sFolderType)
		sInput = InputBox("Please input the " & sFolderType & " folder." & VbCr &_
			VbCr & "To cancel, leave blank or hit the ""Cancel"" button.", sFolderType)
		If sInput = "" Then Wscript.Quit : Else
		If Not FSO.FolderExists(sInput) Then
			Wscript.Echo "The folder you provided (" & sInput & ") does not exist."
			fVerifyFolders(sFolderType)
		End If
		fVerifyFolders = sInput
	End Function
	
	'--------------------------------------------------------------------
	' RecursiveCopy Subroutine (Recursively evaluate subfolders for
	' addtional subfolders and files before performing the copy actions).
	'--------------------------------------------------------------------
	Sub RecursiveCopy(oSrc, oDst)
		'Purge Process
		If iPurge = VbYes Then
			For Each oDstFile in oDst.Files
				If Not FSO.FileExists(oSrc.Path & "\" & oDstFile.Name) Then
					sDstFileSize = oDstFile.Size
					sDstFilePath = oDstFile.Path
					oDstFile.Delete True
					If Err.number <> 0 Then
						oLog.WriteLine(",Purge,X," & fConvertSize(sDstFileSize) & "," & Hex(Err.number) & "," &_
							fTrim(sDest, sDstFilePath))
						iErrFile = iErrFile + 1
					Else
						oLog.WriteLine(",Purge,," & fConvertSize(sDstFileSize) & ",0," & fTrim(sDest, sDstFilePath))
						iPurgeFile = iPurgeFile + 1
					End If
				End If
			Next
			For Each oDstSub in oDst.Subfolders
				If Not FSO.FolderExists(oSrc.Path & "\" & oDstSub.Name) Then
					sDstSubSize = oDstSub.Size
					sDstSubPath = oDstSub.Path
					oDstSub.Delete True
					If Err.number <> 0 Then
						oLog.WriteLine(",Purge,X," & fConvertSize(sDstSubSize) & "," & Hex(Err.number) & "," &_
							fTrim(sDest, sDstSubPath))
						iErrFldr = iErrFldr + 1
					Else
						oLog.WriteLine(",Purge,," & fConvertSize(sDstSubSize) & ",0," & fTrim(sDest, sDstSubPath))
						iPurgeFldr = iPurgeFldr + 1
					End If
				End If
			Next
		End If

		'File Copy/Replace/Skip Process
		For Each oSrcFile in oSrc.Files
			If Not FSO.FileExists(oDst.Path & "\" & oSrcFile.Name) Then
				FSO.CopyFile oSrcFile.Path, oDst.Path & "\", False
				If Err.number <> 0 Then
					oLog.WriteLine(",Copy,X," & fConvertSize(oSrcFile.Size) & "," & Hex(Err.number) & "," &_
						fTrim(sSource, oSrcFile.Path))
					iErrFile = iErrFile + 1
				Else
					oLog.WriteLine(",Copy,," & fConvertSize(oSrcFile.Size) & ",0," & fTrim(sSource, oSrcFile.Path))
					iCopyFile = iCopyFile + 1
				End If
			ElseIf FSO.FileExists(oDst.Path & "\" & oSrcFile.Name) Then
				Set oDstFile = FSO.GetFile(oDst.Path & "\" & oSrcFile.Name)
				If oSrcFile.DateLastModified > oDstFile.DateLastModified Then
					FSO.CopyFile oSrcFile.Path, oDst.Path & "\"
					If Err.number <> 0 Then
						oLog.WriteLine(",Replace,X," & fConvertSize(oSrcFile.Size) & "," & Hex(Err.number) & "," &_
							fTrim(sSource, oSrcFile.Path))
						iErrFile = iErrFile + 1
					Else
						oLog.WriteLine(",Replace,," & fConvertSize(oSrcFile.Size) & ",0," & fTrim(sSource, oSrcFile.Path))
						iReplaceFile = iReplaceFile + 1
					End If
				ElseIf oSrcFile.DateLastModified <= oDstFile.DateLastModified Then
					oLog.WriteLine(",Skip,," & fConvertSize(oSrcFile.Size) & ",0," & fTrim(sSource, oSrcFile.Path))
					iSkipFile = iSkipFile + 1
				End If
			End If
		Next

		'Folder Copy/Replace/Skip Process
		For Each oSrcSub in oSrc.Subfolders
			If Not FSO.FolderExists(oDst.Path & "\" & oSrcSub.Name) Then
				FSO.CopyFolder oSrcSub.Path, oDst.Path & "\", False
				If Err.number <> 0 Then
					oLog.WriteLine(",Copy,X," & fConvertSize(oSrcSub.Size) & "," & Hex(Err.number) & "," &_
						fTrim(sSource, oSrcSub.Path))
					iErrFldr = iErrFldr + 1
				Else
					oLog.WriteLine(",Copy,," & fConvertSize(oSrcSub.Size)& ",0," & fTrim(sSource, oSrcSub.Path))
					iCopyFldr = iCopyFldr + 1
				End If
				Set oDstSub = FSO.GetFolder(oDst.Path & "\" & oSrcSub.Name)
				RecursiveCopy oSrcSub, oDstSub 'Point of Recursion
			ElseIf FSO.FolderExists(oDst.Path & "\" & oSrcSub.Name) Then
				Set oDstSub = FSO.GetFolder(oDst.Path & "\" & oSrcSub.Name)
				If oSrcSub.DateLastModified > oDstSub.DateLastModified Then
					FSO.CopyFolder oSrcSub.Path, oDst.Path
					If Err.number <> 0 Then
						oLog.WriteLine(",Replace,X," & fConvertSize(oSrcSub.Size) & "," & Hex(Err.number) & "," &_
							fTrim(sSource, oSrcSub.Path))
						iErrFldr = iErrFldr + 1
					Else
						oLog.WriteLine(",Replace,," & fConvertSize(oSrcSub.Size) & ",0," & fTrim(sSource, oSrcSub.Path))
						iReplaceFldr = iReplaceFldr + 1
					End If
				ElseIf oSrcSub.DateLastModified <= oDstSub.DateLastModified Then
					oLog.WriteLine(",Skip,," & fConvertSize(oSrcSub.Size) & ",0," & fTrim(sSource, oSrcSub.Path))
					iSkipFldr = iSkipFldr + 1
				End If
				RecursiveCopy oSrcSub, oDstSub 'Point of recursion
			End If
		Next
	End Sub
	
	'--------------------------------------------------------------------
	' fConvertSize Function
	'--------------------------------------------------------------------
	Function fConvertSize(Size) 

		Suffix = " Bytes" 
		If Size >= 1024 Then suffix = " KB" 
		If Size >= 1048576 Then suffix = " MB" 
		If Size >= 1073741824 Then suffix = " GB" 
		If Size >= 1099511627776 Then suffix = " TB" 

		Select Case Suffix 
		    Case " KB" Size = Round(Size / 1024, 2) 
		    Case " MB" Size = Round(Size / 1048576, 2) 
		    Case " GB" Size = Round(Size / 1073741824, 2) 
		    Case " TB" Size = Round(Size / 1099511627776, 2) 
		End Select 

		fConvertSize = Size & Suffix 
	End Function
	
	'--------------------------------------------------------------------
	' fTrim Function (Trim path to remove a portion of the path)
	'--------------------------------------------------------------------
	Function fTrim(RemoveSection, TrimFrom)
	
		TargetLen = Len(TrimFrom)
		RemSection = Len(RemoveSection)
		TrimLen = TargetLen - RemSection
		fTrim = Right(TrimFrom, TrimLen)
	
	End Function

'--------------------------------------------------------------------
' Object CleanUp
'--------------------------------------------------------------------
Set oXl = Nothing
Set oLogRen = Nothing
Set oDstSub = Nothing
Set oSrcSub = Nothing
Set oDstFile = Nothing
Set oSrcFile = Nothing
Set oLog = Nothing
Set oDest = Nothing
Set oSource = Nothing
Set OWS = Nothing
Set FSO = Nothing

'--------------------------------------------------------------------
' Notes on Modifications
'--------------------------------------------------------------------
' Date: 12/20/2011
'   Added Logging.  Tested and confirmed functional.  
'
' Date: 12/21/2011
'   Rearranged the logic so that the Purge Process was executed, then 
'     files were processed, then folders.  
'   Moved all of the main process to the subroutine.  This reduced the
'     lines of code from 316 to 272, thereby making the script more
'     efficient.
'   Added a section to open the log file with excel once the script has
'     completed running.  This will let the executor know that the script
'     has completed.
'   Fixed the fConvertSize function.  Sizes are now logged.
'   Added section to cleanup objects that were used during processing.
'   Added section of Variable Declarations. Activated "Option Explicit".
'   Added a second point of recursion for the RecursiveCopy subroutine.
'     This will allow all items to be logged even if it was already copied.
'
' Date: 2/2/2012
'   Optimized code for the repeating character used as a separator (line 70).
'   
'--------------------------------------------------------------------