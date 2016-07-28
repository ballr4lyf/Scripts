'-------------------------------------------------------------------------------
' Install a local printer using a network port
'
' Created by: Robert D. Rathbun
' Created on: 12/12/2011
'-------------------------------------------------------------------------------

On Error Resume Next

'-------------------------------------------------------------------------------
' Modify This Section to fit individual Printer needs
'-------------------------------------------------------------------------------

sIPAddr = "151.166.212.32"
sPrntDrv = "HP_CLJ_CP3525_32bit_PCL6\hpc3525c.inf"
sPrntName = "413FLTS-CAXI-HP3525"
sModel = "HP Color LaserJet CP3525 PCL 6"

'-------------------------------------------------------------------------------
' Script Body
'-------------------------------------------------------------------------------

Set OWS = Wscript.CreateObject("Wscript.Shell")

'-------------------------------------------------------------------------------
' Create a new Printer Port
'-------------------------------------------------------------------------------

sComputer = "."
Set oWMISrvc = GetObject("winmgmts:" _
  & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
Set oNewPort = oWMISrvc.Get _
  ("Win32_TCPIPPrinterPort").SpawnInstance_

oNewPort.Name = sIPAddr
oNewPort.Protocol = 1
oNewPort.HostAddress = sIPAddr
oNewPort.PortNumber = "9100"
oNewPort.SNMPEnabled = False
oNewPort.Put_

'-------------------------------------------------------------------------------
' Add quotes to the items that will require them.
'-------------------------------------------------------------------------------

sIPAddr = Chr(34) & sIPAddr & Chr(34)
sPrntDrv = Chr(34) & "S:\CAX\Computer Resources\Printers\PrinterDrivers\" &_ 
  sPrntDrv & Chr(34)
sPrntName = Chr(34) & sPrntName & Chr(34)
sModel = Chr(34) & sModel & Chr(34)

'-------------------------------------------------------------------------------
' Run the command to install the printer
'-------------------------------------------------------------------------------

OWS.Run "rundll32 printui.dll PrintUIEntry /q /if /b " & sPrntName & " /f " &_
  sPrntDrv & " /r " & sIPAddr & " /m " & sModel

Set oNewPort = Nothing
Set oWMISrvc = Nothing
Set OWS = Nothing