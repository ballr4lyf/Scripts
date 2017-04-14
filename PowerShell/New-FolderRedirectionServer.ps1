
<#
    Script Name:  New-FolderRedirectionServer
    Created by:  Robert Rathbun
    Created Date:  04/14/2017
    Purpose:  Redirected folders were not being redirected to the new file server.
              The purpose of this script was to fix that issue by updating the appropriate
              registry settings.
#>

$oldServer = "oldServerName"
$newServer = "newServerName"

function replace($key) {
    foreach ($property in (Get-ItemProperty -Path $key.PSPath)) {
        If ($property.toString() -like "\\$oldServer\*") {
            Set-ItemProperty -Path $key.PSPath -Name $property.Name -Value ($property.ToString()).replace($oldServer, $newServer)
        }
    }
}

$shellFolders =Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
$userShellFolders = Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"


# $profileFolders = @("Administrative Tools", "Desktop", "Favorites", "Personal", "Programs", "Start Menu", "Startup")

