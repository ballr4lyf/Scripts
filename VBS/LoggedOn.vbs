On error resume next
strComputer = InputBox("Type in the name of the computer you want to query.")
  If strComputer > "" Then
Set objWMIService = GetObject("winmgmts:" _
  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colComputer = objWMIService.ExecQuery _
   ("Select * from Win32_ComputerSystem")
For Each objComputer in colComputer
Wscript.Echo objComputer.UserName
Next

Else
Wscript.Quit
End If

Wscript.Quit
