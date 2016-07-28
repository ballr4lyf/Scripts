	Option Explicit
	On Error Resume Next

	Dim objShell		: Set objShell = WScript.CreateObject("WScript.Shell")
	Dim objNetwork		: Set objNetwork = WScript.CreateObject("WScript.Network")
	Dim objGroupList
	Dim enumPrinters	: Set enumPrinters = objNetwork.EnumPrinterConnections
	Dim strWorkDir		: strWorkDir = ObjShell.ExpandEnvironmentStrings("%temp%")
	Dim strUser			: strUser = objNetwork.UserName
	Dim strDomain		: strDomain = objNetwork.UserDomain
	Dim strGroup
	Dim objUser			: Set objUser = GetObject("WinNT://" & strDomain & "/" & strUser & ",user")
	Dim intCounter
	Dim localPrinter	: Set localPrinter = False

'Set script working directory to user %temp%

	objShell.CurrentDirectory = strWorkDir

'Map standard printers for all users

'If local printer exists on LPT or USB, set to default

	For intCounter = 0 to enumPrinters.Count -1 step 2
		If Left(enumPrinters(intCounter),3)="LPT" OR Left(enumPrinters(intCounter),3)="USB" OR Left(enumPrinters(intCounter),3)="DOT" Then
			If Left(enumPrinters(intCounter+1),7)="Acrobat" Then
			Else
			objNetwork.SetDefaultPrinter enumPrinters(intCounter+1)
			localPrinter = True
			End If
		End If
	Next

'Map additional printers and change default printer If no local printer based on group membership

	' The following security groups can be found in the OU: "EGLIN-2K\TEAS\AGM130 - Hanger workstations"
	
	strGroup = "AGM-130_HP_CLJ4005_PS"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\AGM-130_HP_CLJ4005_PS"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\AGM-130_HP_CLJ4005_PS"
	End If

	
'Cleanup

	Set objGroupList = Nothing
	Set objUser = Nothing

'Function to test group membership

	Function IsMember(strGroup)

	If IsEmpty(objGroupList) Then
	Call LoadGroups
	End If

	IsMember = objGroupList.Exists(strGroup)

	End Function

'Subroutine to load user's groups into dictionary object

	Sub LoadGroups

	Dim objGroup

	Set objGroupList = CreateObject("Scripting.Dictionary")
	objGroupList.CompareMode = vbTextCompare
	For Each objGroup In objUser.Groups
	objGroupList(objGroup.name) = True
	Next

	Set objGroup = Nothing

	End Sub