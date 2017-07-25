<# 
    Script Name:  Install-DattoAgent
    Created By:  Robert D. Rathbun
    Created Date:  08/01/2016
    Purpose:  Download and silently install the Datto Windows Agent.

    Note:
    1.  Invoke-WebRequest requires PowerShell v3 or higher.
    2.  The PC/Server will still need to be rebooted after install.
#>

$URL = "https://www.datto.com/downloads/DattoWindowsAgent.exe"
$installDir = "C:\Kits"
$installer = $installDir + "\DattoWindowsAgent.exe"


# Check PowerShell version is greater than or equal to 3.
If ($PSVersionTable.PSVersion.Major -ge 3) {

    # Check for folder path and install file.
    If (!(Test-Path $installDir)) {
        New-Item $installDir -ItemType Directory
    } ElseIf (Test-Path $installer) {
        Remove-Item -Force $installer
    }

    Invoke-WebRequest -Uri $URL -OutFile $installer #Download the Datto Windows Agent.

    Start-Process $installer -ArgumentList "/S"  #Silent install.
} Else {
    Write-Output "Script requires PowerShell version 3 or higher."
    Exit
}
