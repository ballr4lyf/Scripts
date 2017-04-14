
<#
    Created by:  Robert Rathbun
    Created Date:  01/11/2016
    Purpose:  Redirected folders were not being redirected to the new file server.
              The purpose of this script was to fix that issue by updating the appropriate
              registry settings.
#>

$shellFolders = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
$userShellFolders = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

$oldServer = "oldServerName"
$newServer = "newServerName"

# $profileFolders = @("Administrative Tools", "Desktop", "Favorites", "Personal", "Programs", "Start Menu", "Startup")

