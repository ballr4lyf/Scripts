# Enable PSRemoting remotely without using PSExec.

$ComputerName = "MyPC01"

([wmiclass]"\\$ComputerName\root\cimv2:win32_process").create('powershell "Enable-PsRemoting -Force"')
