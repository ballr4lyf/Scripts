Set OWS = Wscript.CreatObject("Wscript.Shell")
Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")

Set colPrinters = oWMI.ExecQuery("SELECT * FROM Win32_Printer")

sOldServer = "SRV-01"
arrOldPrinters = Array("KM-252-C", _
					   "KM-C280-C", _
					   "KM-C252-BW", _
					   "KM-C280-BW", _
					   "KM-C280-FAX")

For Each sOldPrinter in arrOldPrinters
	For Each oPrinter in colPrinters
		If UCASE(oPrinter.DeviceID) = UCASE("\\" & sOldServer & "\" & sOldPrinter) Then
			OWS.Run("rundll32 printui.dll PrintUIEntry /dn /q /n " & """\\" & sOldServer & "\" & sOldPrinter & """", 1, true)
		End If
	Next
Next