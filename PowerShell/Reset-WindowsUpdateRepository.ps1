Stop-Service -Name BITS -Force
Stop-Service -Name wuauserv -Force

If (Test-Path C:\Windows\SoftwareDistribution.old){
    Remove-Item -Path C:\Windows\SoftwareDistribution.old -Recurse -Force
}

Rename-Item -Path C:\Windows\SoftwareDistribution -NewName SoftwareDistribution.old

Start-Service -Name BITS
Start-Service -Name wuauserv 