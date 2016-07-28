'--------------------------------------------------------------------------
' Delete Files older than "X" days old.
' 
' Created By: Robert D. Rathbun
' Created Date: 12/8/2011
' Modified Date: 
' 
'--------------------------------------------------------------------------

Option Explicit

Dim FSO, oFolder, oFile, oSubFolder
Dim iStart, iCont, iSubFolder
Dim sFolder, sDaysOld

Set FSO = CreateObject("Scripting.FileSystemObject")

iStart = MsgBox("This script will delete files and subfolders older than a specified number " &_
  "of days."  & VbCrLF & VbCrLf & "Do you wish to continue?", VbYesNo, "Delete Old Files")
  
  If iStart = VbYes Then

    DeclareFolder
    DeclareDaysOld
    
    Set oFolder = FSO.GetFolder(sFolder)
      
      For Each oFile in oFolder.Files
        If oFile.DateLastModified < (Date() - sDaysOld) Then
          ofile.Delete(True)
        End If
      Next
      
      iSubFolder = MsgBox("Are you sure you want to delete all subfolders older than " & sDaysOld &_
        " days old?  All files in the subfolders will be deleted.", VbYesNo, "Delete Subfolders")
        
      If iSubFolder = VbYes Then
      
        For Each oSubfolder in oFolder.SubFolders
          If oSubFolder.DateLastModified < (Date() - sDaysOld) Then
            oSubFolder.Delete(True)
          End If
        Next
        
    Else
    
    CreateObject("Wscript.Shell").Popup "Script Cancelled.", 5, "Cancelled"
    
    End If
  
  Else
  
  CreateObject("Wscript.Shell").Popup "Script Cancelled.", 5, "Cancelled"
  
  End If

'--------------------------------------------------------------------------
' Declare the folder to be examined
'--------------------------------------------------------------------------
Sub DeclareFolder
  sFolder = InputBox("Please input the root folder from which you want to delete old " &_
    "files.", "Root Folder")
    If Not FSO.FolderExists(sfolder) Then
      Wscript.Echo "The Folder Provided " & Chr(34) & sFolder & Chr(34) & " does not exist." &_
        VbCrLf & VbCrLf & "Please re-enter the root folder you wish to delete old files/folders from."
      DeclareFolder
    End If
End Sub

'--------------------------------------------------------------------------
' Declare the number of days old
'--------------------------------------------------------------------------
Sub DeclareDaysOld
  sDaysOld = InputBox("Please enter the age of files/folders to delete (in days).", "Days Old")
    If Not IsNumeric(sDaysOld) Then
      Wscript.Echo "The amount you entered is not numeric." & VbCrLf & VbCrLf & "Please enter a numeric " &_
        "value for the number of days old."
      DeclareDaysOld
    End If
    sDaysOld = CInt(sDaysOld)
End Sub