<#
    Script Name:  Restart-DattoServices.ps1
    Author:  Rob Rathbun
    Date:  07/11/2017
    Purpose:  Simple script to fix issues with ShadowSnap services for Datto agents.
#>

$services = @("vsnapvss", "ShadowProtectSvc", "stc_raw_agent")

foreach ($service in $services) {
    Stop-Service -Name $service
    If (Get-Process -Name $service -ErrorAction SilentlyContinue) {
        Stop-Process -Name $service -Force
    }
    Start-Service -Name $service
}