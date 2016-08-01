<# 
    Script Name:  Install-DattoAgent
    Created By:  Robert D. Rathbun
    Created Date:  08/01/2016
    Purpose:  Download and silently install the Datto Windows Agent.

    Note:
    1.  Invoke-WebRequest requires PowerShell v3 or higher.
#>

$URL = "http://downloads.dattobackup.com/ShadowSnap/DattoWindowsAgent51.exe"
$installDir = "C:\Kits"
$installer = $installDir + "\DattoWindowsAgent51.exe"

If (!(Test-Path $installDir)) {
    New-Item $installDir -ItemType Directory
} ElseIf (Test-Path $installer) {
    Remove-Item -Force $installer
}

Invoke-WebRequest -Uri $URL -OutFile $installer #Download the Datto Windows Agent.

Start-Process $installer -ArgumentList "/S"  #Silent install.
