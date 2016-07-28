'====================================================================
'	Compare contents of folder to contents of backups.  Log only errors.
'
'	Created By:  Robert D. Rathbun
'	Created Date:  3/9/2012
'	Modified Date:
'
'====================================================================

On Error Resume Next

Dim FSO, OWS, oSource, iSourceLen, sDest, sLog, oLog

Set FSO = CreateObject("Scripting.FileSystemObject")
Set OWS = CreateObject("Wscript.Shell")

Set oSource = FSO.GetFolder("S:\")
'Set oDest = FSO.GetFolder("E:\413FLTS Common")

iSourceLen = Len(oSource.Path)
sDest = "E:\413FLTS Common\"

sLog = "E:\Logs\File Compare" & Year(Now) & "-" & fDblDigit(Month(Now)) & "-" & fDblDigit(Day(Now)) & "_FileCompare.csv"

If FSO.FileExists(sLog) Then
	Set oLog = FSO.OpenTextFile(sLog, 8)		' ForAppending = 8
	oLog.WriteBlankLines(1)
	oLog.WriteLine(String(150, "-"))
Else
	FSO.CreateTextFile(sLog)
End If

oLog.WriteLine("Comfirm that contents of the source were copied to the Backup Directory.")
oLog.WriteLine("NOTE: Only items logged are those that were NOT copied to the Backup Directory.")
oLog.WriteLine("Source:  " & oSource.Path)
oLog.WriteLine("Backup Directory:  " & sDest)
oLog.WriteLine("Comparison Started:  " & Now)
oLog.WriteBlankLines(1)
oLog.WriteLine(",File Size (MB),Source Path")

CompareFiles oSource

oLog.WriteBlankLines(1)
oLog.WriteLine("Comparison Ended:  " & Now)

'--------------------------------------------------------------------
' Functions and Subroutines
'--------------------------------------------------------------------
	'--------------------------------------------------------------------
	' Compare Files Subroutine
	'--------------------------------------------------------------------
	Sub CompareFiles(Source)
		Dim oFile, oSub, sCompareString
		
		For Each oFile in Source.Files
			sCompareString = sDest & Right(oFile.Path, oFile.Path - iSourceLen)
			If Not FSO.FileExists(sCompareString) Then
				oLog.WriteLine("," & fConvertSize(oFile.Size) & "," & fCommaCheck(oFile.Path))
			End If
		Next
		
		For Each oSub in Source.SubFolders
			CompareFiles(oSub)
		Next
	End Sub
	
	'--------------------------------------------------------------------
	' fDblDigit Function
	'--------------------------------------------------------------------
	Function fDblDigit(Number)
		Select Case Number
			Case 1, 2, 3, 4, 5, 6, 7, 8, 9
				fDblDigit = "0" & Number
			Case Else
				fDblDigit = Number
		End Select
	End Function
	
	'--------------------------------------------------------------------
	' fConvertSize Function
	'--------------------------------------------------------------------
	Function fConvertSize(Size)
		Size = Round(Size / 1048576, 3)

		fConvertSize = Size
    End Function
	
	'--------------------------------------------------------------------
	' fCommaCheck Function
	'--------------------------------------------------------------------
	Function fCommaCheck(Path)
		fCommaCheck = Replace(Path, ",", ";")
    End Function