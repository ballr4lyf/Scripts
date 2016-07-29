
<#
    Created By:  Robert Rathbun
    Created Date:  07/29/2016
    Purpose:  Remove corrupted profile registry key on a PC.
#>

Set-Location HKLM:\

Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object {
    If ($_.Name -like "*.bak") {
        Remove-Item -Path $_.Name
    }
}
