
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

function replaceServer($key) {
    $properties = (Get-ItemProperty -Path $key.PSPath)
    $a = $properties | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name
    foreach ($propertyName in $a) {
        If ($properties.$propertyName.ToString() -like "\\$oldServer*") {
            $newValue = $properties.$propertyName.ToString() -replace $oldServer, $newServer
            Set-ItemProperty -Path $key.PSPath -Name $propertyName -Value $newValue
        }
    }
}

$shellFolders =Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
$userShellFolders = Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

replaceServer($shellFolders)
replaceServer($userShellFolders)

Stop-Process -Name explorer
Start-Process -FilePath C:\Windows\explorer.exe