'====================================================================
'   Script to crawl a directory, search for old format MS Office 
'   documents (.doc, .xls, and .ppt), and convert them to newer format
'   MS Office documents (.docx, .xlsx, and .pptx).
'
'	Created by: Robert D. Rathbun
'	Created Date: 2/7/2012
'	Modified Date: 2/15/2012
'
'====================================================================

Option Explicit
'On Error Resume Next

Dim OWS, FSO, oLog, oFolder, sFolder, sArchiveFolder, sWorkingFolder

Set OWS = CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

sArchiveFolder = "C:\Archive"

If Not FSO.FolderExists(sArchiveFolder) Then
	FSO.CreateFolder(sArchiveFolder)
End If

If Not FSO.FolderExists(sArchiveFolder & "\_Logs") Then
	FSO.CreateFolder(sArchiveFolder & "\_Logs")
End If

'Open or Create the log file.
If FSO.FileExists(sArchiveFolder & "\_Logs\MSOffice_File_Conversion.csv") Then
	Set oLog = FSO.OpenTextFile(sArchiveFolder & "\_Logs\MSOffice_File_Conversion.csv", 8)
	oLog.WriteBlankLines(1)
	oLog.WriteLine(String(150, "-"))
Else
	Set oLog = FSO.CreateTextFile(sArchiveFolder & "\_Logs\MSOffice_File_Conversion.csv")
End If

'Obtain the folder to examine
sFolder = fGetFolder
Set oFolder = FSO.GetFolder(sFolder)

'Write the header lines of the log file.
oLog.WriteLine("Folder Examined: " & oFolder.Path)
oLog.WriteLine("Started: " & Now)
oLog.WriteLine("All sizes are in Megabytes. (NOTE: 1024 MegaBytes = 1 GigaByte)")
oLog.WriteLine(",Type,Original Size (MB),Converted Size(MB),Document Path")

'Perform the main function of the script
ExamineFolder oFolder

'Ending Lines of the log file
oLog.WriteBlankLines(1)
oLog.WriteLine("Finished: " & Now)

Wscript.Echo "Examination of folder""" & oFolder.Path & """ completed."

'--------------------------------------------------------------------
' Functions and Subroutines
'--------------------------------------------------------------------
	'--------------------------------------------------------------------
	' fVerifyFolder Function
	'--------------------------------------------------------------------
	Function fGetFolder
		Dim sInput
	
		'Obtain directory information from user and verify that directory exists
		sInput = InputBox("Please input the paty to the folder you wish to examine." &_
			VbCr & VbCr & "To cancel, leave blank or hit the ""Cancel"" button.", "Folder")
		If sInput = "" Then Wscript.Quit
		If Not FSO.FolderExists(sInput) Then
			Wscript.Echo "The folder you provided (" & sInput & ") does not exist."
			fGetFolder
		Else
			fGetFolder = sInput
		End If
	
	End Function
	
	'--------------------------------------------------------------------
	' ExamineFolder subroutine
	'--------------------------------------------------------------------
	Sub ExamineFolder(Folder)
		Dim sExt, oFile, oSub
		Dim sType, sFileName, sFilePath, sOldSize, sNewSize, oApp, oDoc, oArchivedFile
		Dim iNewFileExt, sNewFile, oConvertedFile
		
		sWorkingFolder = Right(Folder.Path, Len(Folder.Path) - 2)
		
		If Not FSO.FolderExists(sArchiveFolder & sWorkingFolder) Then
			FSO.CreateFolder(sArchiveFolder & sWorkingFolder)
		End If
		
		sWorkingFolder = sArchiveFolder & sWorkingFolder & "\"
		
		For Each oFile in Folder.Files
			sExt = "." & LCase(FSO.GetExtensionName(oFile.Path))
			sOldSize = oFile.Size
			sFileName = Left(oFile.Name, Len(oFile.Name) -4)
			sFilePath = Left(oFile.Path, Len(oFile.Path) - Len(oFile.Name))
			
			'Exit loop if file is a temporary file.
			If Left(sFileName, 2) = "~$" Then Exit For
			
			sNewFile = sWorkingFolder & sFileName & sExt
			
			Select Case sExt
				Case ".doc"
					oFile.Copy sNewfile, True
					sType = "WORD"
					Set oApp = CreateObject("Word.Application")
					Set oDoc = oApp.Application.Documents.Open(sNewFile)
					iNewFileExt = 12
					
				Case ".xls"
					oFile.Copy sNewfile, True
					sType = "EXCEL"
					Set oApp = CreateObject("Excel.Application")
					Set oDoc = oApp.Application.Workbooks.Open(sNewFile)
					iNewFileExt = 51
					
				Case ".ppt"
					oFile.Copy sNewfile, True
					sType = "POWERPOINT"
					Set oApp = CreateObject("PowerPoint.Application")
					oApp.Visible = True 'Required option for PowerPoint
					Set oDoc = oApp.Presentations.Open(sNewFile)
					
				Case Else
					Exit For
					
			End Select
			
			'Check if the file has already been converted
			If FSO.FileExists(sFilePath & sFileName & sExt & "x") Then
				Set oConvertedFile = FSO.GetFile(sFilePath & sFileName & sExt & "x")
				
				'Check if an Archived copy of the original file exists.
				If FSO.FileExists(sWorkingFolder & sFileName & sExt) Then
					Set oArchivedFile = FSO.GetFile(sWorkingFolder & sFileName & sExt)
					
					'Confirm that the converted copy is the newest of all files.
					If oArchivedFile.DateLastModified > oFile.DateLastModified AND _
					oConvertedFile.DateLastModified > oFile.DateLastModified Then
						sNewSize = oConvertedFile.Size
						oFile.Copy sNewFile, True
						oFile.Delete
						oLog.WriteLine("," & sType & "," & fConvertSize(sOldSize) & "," &_
							fConvertSize(sNewSize) & "," & fCommaCheck(oConvertedFile.Path))
						Set oConvertedFile = Nothing
						Set oArchivedFile = Nothing
						Exit For
					End If
				End If
			End If
		
			'Convert the Archived file (does not retain the original document)
			oDoc.SaveAs sNewFile & "x", iNewFileExt
			oDoc.Close
			oApp.Quit
			
			'After conversion, make another copy of the original document (for backup purposes)
			oFile.Copy sNewFile, True
			sNewFile = sNewFile & "x"
			
			'Cleanup object variables for reuse on next file.
			Set oDoc = Nothing
			Set oApp = Nothing
			
			'Delete the original file, and copy the newly converted file to original location.
			oFile.Delete
			FSO.CopyFile sNewFile, sFilePath, True
			
			Set oConvertedFile = FSO.GetFile(sFilePath & sFileName & sExt & "x")
			sNewSize = oConvertedFile.Size
			
			'Write a line to the log file
			oLog.WriteLine("," & sType & "," & fConvertSize(sOldSize) & "," &_
				fConvertSize(sNewSize) & "," & fCommaCheck(oConvertedFile.Size))
			Set oConvertedFile = Nothing
		Next
			
		'Crawl through each subfolder in the root directory
		For Each oSub in Folder.Subfolders
			ExamineFolder oSub
		Next
	
	End Sub
	
	'--------------------------------------------------------------------
    ' Function to convert folder and file sizes to Megabytes.
    '--------------------------------------------------------------------
    Function fConvertSize(Size)
    
    Size = Round(Size / 1048576, 3)
    fConvertSize = Size

    End Function
    
    '--------------------------------------------------------------------
    ' Function to Check for commas within a string.
    '--------------------------------------------------------------------
    Function fCommaCheck(Path)
		fCommaCheck = Replace(Path, ",", ";")
    End Function

' Changes from v2 to v3
'	1. Renamed the main subroutine
'	2. Added more concise notes to the main subroutine.
'	3. Located and fixed issue that was preventing the Script from completing.
'	4. Optimized the code.  Reduced the number of code lines by 40+ lines.
'	5. Changed the working folder to mirror the original folder path.