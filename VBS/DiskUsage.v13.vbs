'====================================================================
'
'Generate a report to get File and Folder sizes on a share.  The report
'is created on the user's desktop.
'
'Created by: Robert D. Rathbun
'Created Date: 11/21/2011
'Last Modified Date:  1/4/2012
'
'====================================================================

'--------------------------------------------------------------------
'Requires that variables be declared using the "Dim" statement before
' they are used.
'--------------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------------
'Variable Declaration
'--------------------------------------------------------------------
Dim OWS, FSO, oGetFile, oTextFile, oFolder, oFile, oReportFolder, oSubFolder, oXl, oXlBook 'Objects
Dim sUserProfile, sFolder, sLogFile, sTargetFile, sTempFolder, sFileX 'Strings

'--------------------------------------------------------------------
'If an error occurs, proceed to the next line
'--------------------------------------------------------------------
On Error Resume Next

Set OWS = Wscript.CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

'--------------------------------------------------------------------
'Enumerate the USERPROFILE Environment Variable (standard in Windows)
'  and the TEMP Environment Variable (also standard in Windows)
'--------------------------------------------------------------------
sUserProfile = OWS.ExpandEnvironmentStrings("%USERPROFILE%")
sTempFolder = OWS.ExpandEnvironmentStrings("%TEMP%")

'--------------------------------------------------------------------
'Call to the "VerifyFolderPath" Subroutine.
'--------------------------------------------------------------------
VerifyFolderPath

'--------------------------------------------------------------------
'Provide a name for the Log File
'--------------------------------------------------------------------
sLogFile = Right(sFolder, Len(sFolder) - 3) & "_DiskUsageReport"
	If FSO.FileExists(sTempFolder & "\" & sLogFile & ".csv") Then
	  FSO.DeleteFile(sTempFolder & "\" & sLogFile & ".csv")
	End If
	Set oTextFile = FSO.CreateTextFile(sTempFolder & "\" & sLogFile & ".csv")

'--------------------------------------------------------------------
'Sets the folder as an Object
'--------------------------------------------------------------------
Set oFolder = FSO.GetFolder(sFolder)

'--------------------------------------------------------------------
'Writing the header lines of the LogFile
'--------------------------------------------------------------------
oTextFile.WriteLine("Folder Examined: " & sFolder)
oTextFile.WriteLine("Started: " & Now)
oTextFile.WriteLine("All Sizes are in MegaBytes. (NOTE: 1024 MegaBytes = 1 GigaByte)")
oTextFile.WriteBlankLines(1)
oTextFile.WriteLine(",Folder Size(MB),File Size(MB),File Type,Date Last Modified,Path")
oTextFile.WriteLine("," & fConvertSize(oFolder.Size) & ",,," &_
	FormatDateTime(oFolder.DateLastModified, 2) & "," & fCommaCheck(oFolder.Path))

'--------------------------------------------------------------------
'Call to the "EnumSubFolders" Subroutine (Main Process)
'--------------------------------------------------------------------
EnumSubFolders oFolder

'--------------------------------------------------------------------
'Write the closing lines of the Log File and close the file.
'--------------------------------------------------------------------
oTextFile.WriteBlankLines(1)
oTextFile.WriteLine("Ended: " & Now)

oTextFile.Close

'--------------------------------------------------------------------
'Convert the Log File file to *.xlsx to make it easier to sort
'and read.
'--------------------------------------------------------------------
Set oReportFolder = FSO.GetFolder(sTempFolder)
sTargetFile = sUserProfile & "\desktop\" & sLogFile & ".xlsx"
	If FSO.FileExists(sTargetFile) Then
		FSO.DeleteFile(sTargetFile)
	End If

Set oXl = CreateObject("Excel.Application")
Set	oXlBook = oXl.Application.Workbooks.Open(sTempFolder & "\" & sLogFile & ".csv")
	
	oXl.Application.Visible = False
	oXl.Application.DisplayAlerts = False
	
	oXLBook.SaveAs sTargetFile, 51
	oXLBook.Close
	
	oXl.Application.Visible = True
	oXl.Application.DisplayAlerts = True
	oXl.Application.Workbooks.Open(sTargetFile)
	
	If FSO.FileExists(sTempFolder & "\" & sLogFile & ".csv") Then
		FSO.DeleteFile(sTempFolder & "\" & sLogFile & ".csv")
	End If
	
CreateObject("Wscript.Shell").Popup "Examination of " & sFolder & " completed.", _
	10, "Complete"

'--------------------------------------------------------------------
'Subroutines and Functions:
'--------------------------------------------------------------------

    '--------------------------------------------------------------------
    'Subroutine to recursively enumerate subfolders
    '--------------------------------------------------------------------
    Sub EnumSubFolders(oFolder)
		For Each oFile in oFolder.Files
			oTextfile.WriteLine(",," & fConvertSize(oFile.Size) & "," &_
				fFileX(oFile.Path) & "," & FormatDateTime(oFile.DateLastModified, 2) &_
				"," & fCommaCheck(oFile.Path))
		Next
		For Each oSubfolder in oFolder.Subfolders
			oTextFile.WriteLine("," & fConvertSize(oSubfolder.Size) & ",,," &_
			    FormatDateTime(oSubfolder.DateLastModified, 2) & "," &_
			    fCommaCheck(oSubfolder.Path))
			EnumSubFolders oSubFolder
		Next
	End Sub

    '--------------------------------------------------------------------
    'Subroutine to verify that the provided folder path is a valid path.
    '--------------------------------------------------------------------
    Sub VerifyFolderPath
		sFolder = InputBox("Please input the folder path you wish to examine." &_
			VbCrLf & VbCrLF & "An Excel report will be generated and placed on your desktop.  " &_
			"You will be notified upon completion via popup." & VbCR & VbCr & "To cancel hit ""Cancel"", or " &_
			"just leave blank.")
			If sFolder = "" Then Wscript.Quit : Else
			If Not FSO.FolderExists(sFolder) Then
				Wscript.Echo "Folder Location not Found." & VbCrLf & VbCrLf &_
				"Please input a valid folder Path"
				VerifyFolderPath
		    End If
    End Sub

    '--------------------------------------------------------------------
    'Function to convert folder and file sizes to Megabytes to make the
    '  report more readable.
    '--------------------------------------------------------------------
    Function fConvertSize(Size)
    Size = Round(Size / 1048576, 3)

    fConvertSize = Size
    End Function

    '--------------------------------------------------------------------
    'Function to remove commas (,) from file and folder path names.
    '--------------------------------------------------------------------
    Function fCommaCheck(Path)
		fCommaCheck = Replace(Path, ",", ";")
    End Function
    
    '--------------------------------------------------------------------
    'Function to identify common file extensions.
    '--------------------------------------------------------------------
    Function fFileX(sPath)
		sFileX = LCase(FSO.GetExtensionName(sPath))
		Select Case sFileX
		Case "doc"
			fFileX = "MS Word (.doc)"
		Case "docx"
			fFileX = "MS Word (.docx)"
		Case "log"
			fFileX = "Log File (.log)"
		Case "msg"
			fFileX = "Outlook Message (.msg)"
		Case "txt"
			fFileX = "Text (.txt)"
		Case "ppt"
			fFileX = "MS PowerPoint (.ppt)"
		Case "pptx"
			fFileX = "MS PowerPoint (.pptx)"
		Case "m3u"
			fFileX = "Audio (.m3u)"
		Case "m4a"
			fFileX = "Audio (.m4a)"
		Case "mp3"
			fFileX = "Audio (.mp3)"
		Case "wav"
			FfileX = "Audio (.wav)"
		Case "wma"
			fFileX = "Audio (.wma)"
		Case "avi"
			fFileX = "Video (.avi)"
		Case "mov"
			fFileX = "Video (.mov)"
		Case "mp4"
			fFileX = "Video (.mp4)"
		Case "mpg"
			fFileX = "Video (.mpg)"
		Case "mpeg"
			fFileX = "Video (.mpeg)"
		Case "wmv"
			fFileX = "Video (.wmv)"
		Case "bmp"
			fFileX = "Picture (.bmp)"
		Case "gif"
			fFileX = "Picture (.gif)"
		Case "jpg"
			fFileX = "Picture (.jpg)"
		Case "jpeg"
			fFileX = "Picture (.jpeg)"
		Case "tif"
			fFileX = "Picture (.tif)"
		Case "xls"
			fFileX = "MS Excel (.xls)"
		Case "xlsx"
			fFileX = "MS Excel (.xlsx)"
		Case "xlsm"
			fFileX = "MS Excel (.xlsm)"
		Case "accdb"
			fFileX = "MS Access Database (.accdb)"
		Case "pst"
			fFileX = "MS Outlook Personal Folders (.pst)"
		Case "dwg"
			fFileX = "AutoCAD Drawing (.dwg)"
		Case "iso"
			fFileX = "Disc Image (.iso)"
		Case "bak"
			fFileX = "Backup File (.bak)"
		Case "bkf"
			fFileX = "Backup File (.bkf)"
		Case "tmp"
			fFileX = "Temp File (.tmp)"
		Case "torrent"
			fFileX = "BitTorrent File (.torrent)"
		Case "vsd"
			fFileX = "MS Visio Drawing (.vsd)"
		Case "mpp"
			fFileX = "MS Project (.mpp)"
		Case Else
			fFileX = sFileX
		End Select
    End Function

'--------------------------------------------------------------------
'Cleanup Variables
'--------------------------------------------------------------------
Set oXlBook = Nothing
Set oXl = Nothing
Set oFolder = Nothing
Set oTextFile = Nothing
Set oGetFile = Nothing
Set FSO = Nothing
Set OWS = Nothing

'====================================================================
'Notes on Modifications:
'
'Date:  11/22/2011
'  Added popup at the end of the script to notify that the process 
'    completed.
'  Added the option for the user to specify the report name.
'  Added the function to convert the size from bytes to more readable
'    sizes.
'  Change the created text file to a hidden file (not working).  Moved 
'    the file to the root of the user's profile.
'  Added the "Date Last Modified" attribute to the list of attributes
'    retrieved.
'
' Date: 11/23/2011
'  Added a function to remove commas from the file and folder path
'    names.
'
' Date:  11/28/2011
'   Declared variables and added the "Option Explicit" line.
'   Removed section that was supposed to "hide" the text file, as it was
'     not working.
'   Added "On Error Resume Next" line.
'
' Date:  11/29/2011
'   Added additional comment lines to help troubleshoot the script in the
'     future.
'   Added subroutine to verify that the folder path is a valid path.
'
' Date:  12/12/2011
'   Move the text file created to the "Temp" folder, as defined by the OS
'     environment string.
'   Insert a check to delete the txt file if it exists at the temp directory.
'
' Date:  12/15/2011
'   Changed the header lines to reflect that sizes are listed in MB.
'
' Date: 12/27/2011
'   Set the log file to write directly to a .csv file instead of a .txt
'      file that was then converted to .csv.
'   Removed the section that checked for .txt files to conver to .csv.
'   Added a section that converted the .csv file to .xlsx.
'   Re-organized the openning lines of the log file so that the start time
'      was written after the examined folder is identified.
'
' Date: 12/28/2011
'   Added tracking of file extensions.
'   Re-arranged logic so that all of the main process is done at the subroutine.
'   Removed unneeded variables (colFiles).
'
' Date: 1/4/2012
'   Added function to identify most commonly used file extension.
'
' Date: 2/21/2012
'   No Longer requires input for the "Log File" name.
'====================================================================