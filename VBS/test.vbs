Set oShell = CreateObject("WScript.Shell")
Set getOSVersion = oShell.exec("%comspec% /c ver")
sVersion = getOSVersion.stdout.readall
Wscript.echo sVersion

Select Case True
	Case InStr(sVersion, "n 5.") > 1 : sGetOS = "XP"
	Case InStr(sVersion, "n 6.") > 1 : sGetOS = "Vista"
	Case Else : sGetOS = "Unknown"
End Select

WScript.Echo sGetOS