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

	'For intCounter = 0 to enumPrinters.Count -1 step 2
		'If Left(enumPrinters(intCounter),3)="LPT" OR Left(enumPrinters(intCounter),3)="USB" OR Left(enumPrinters(intCounter),3)="DOT" Then
			'If Left(enumPrinters(intCounter+1),7)="Acrobat" Then
			'Else
			'objNetwork.SetDefaultPrinter enumPrinters(intCounter+1)
			'localPrinter = True
			'End If
		'End If
	'Next

'Map additional printers and change default printer If no local printer based on group membership

	' The following security groups can be found in the OU: "EGLIN-2K\TEAS\AQ-GROUPS\PRINTER DEPLOYMENT"
	
	strGroup = "AQM_Canon_iR_C3480_UFRII"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\AQM_Canon_iR_C3480_UFRII"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\AQM_Canon_iR_C3480_UFRII"
	End If

	strGroup = "AQM_HP_CP2025dn_PCL"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\AQM_HP_CP2025dn_PCL"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\AQM_HP_CP2025dn_PCL"
	End If

	strGroup = "AQM_HP_CP2025dn_PS"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\AQM_HP_CP2025dn_PS"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\AQM_HP_CP2025dn_PS"
	End If	

	strGroup = "AQO_HP_CP2025dn_PCL"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\AQO_HP_CP2025dn_PCL"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\AQO_HP_CP2025dn_PCL"
	End If

	strGroup = "AQO_HP_CP2025dn_PS"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\AQO_HP_CP2025dn_PS"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\AQO_HP_CP2025dn_PS"
	End If	

	strGroup = "B11FL2_HP3800_ENFO"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B11FL2_HP3800_ENFO"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B11FL2_HP3800_ENFO"
	End If

	strGroup = "Bldg11_floor2_enqt"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\Bldg11_floor2_enqt"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\Bldg11_floor2_enqt"
	End If	

	strGroup = "ENL-Canon iR C3220"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\ENL-Canon iR C3220"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\ENL-Canon iR C3220"
	End If

	strGroup = "ENL-HPCLJ4650"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\ENL-HPCLJ4650"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\ENL-HPCLJ4650"
	End If

	strGroup = "ENL-HPCLJ4700"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\ENL-HPCLJ4700"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\ENL-HPCLJ4700"
	End If

	strGroup = "ENQ-Canon IR C3480"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\ENQ-Canon IR C3480"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\ENQ-Canon IR C3480"
	End If

	strGroup = "ENQT-HPLJ4050"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\ENQT-HPLJ4050"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\ENQT-HPLJ4050"
	End If

	strGroup = "TRAPPRGM"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\TRAPPRGM"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\TRAPPRGM"
	End If
	
	' The following security groups can be found in the OU: "EGLIN-2K\TEAS\DAU-GROUPS\PRINTER DEPLOYMENT"

	strGroup = "DAU_Classroom1"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\DAU_Classroom1"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\DAU_Classroom1"
	End If

	strGroup = "DAU_Classroom2"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\DAU_Classroom2"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\DAU_Classroom2"
	End If

	strGroup = "DAUMULTI"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\DAUMULTI"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\DAUMULTI"
	End If

	' The following security groups can be found in the OU: "EGLIN-2K\TEAS\EN-GROUPS\PRINTER DEPLOYMENT"

	strGroup = "B11c1108"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B11c1108"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B11c1108"
	End If

	strGroup = "B11C1109"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B11C1109"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B11C1109"
	End If

	strGroup = "B11R152_TEAS_HPLJ3700N"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B11R152_TEAS_HPLJ3700N"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B11R152_TEAS_HPLJ3700N"
	End If

	strGroup = "B11R160"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B11R160"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B11R160"
	End If

	strGroup = "B349C225A"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B349C225A"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B349C225A"
	End If

	strGroup = "B349C225A_PS"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B349C225A_PS"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B349C225A_PS"
	End If

	strGroup = "B349C225B_PS"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B349C225B_PS"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B349C225B_PS"
	End If

	strGroup = "B349C225C"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B349C225C"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B349C225C"
	End If

	strGroup = "B349C225C_PS"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B349C225C_PS"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B349C225C_PS"
	End If

	strGroup = "B349C225D"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B349C225D"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B349C225D"
	End If

	strGroup = "B349C225D_PS"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\B349C225D_PS"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\B349C225D_PS"
	End If

	strGroup = "ENR-C5035"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\ENR-C5035"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\ENR-C5035"
	End If

	strGroup = "ENRO-HPLJ4050"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\ENRO-HPLJ4050"
		objNetwork.AddWindowsPrinterConnection "\\en-prt\ENRO-HPLJ4050"
	End If

	strGroup = "HP LaserJet 4250 PCL 6"
	If IsMember(strGroup) Then
		Wscript.Echo "Checking for Printer: \\EN-PRT\HP LaserJet 4250 PCL 6"
			objNetwork.AddWindowsPrinterConnection "\\en-prt\HP LaserJet 4250 PCL 6"
		'objWshShell.Run "rundll32 printui.dll,PrintUIEntry /in /Gw /q /n ""\\en-prt\HP LaserJet 4250 PCL 6""",1, true
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