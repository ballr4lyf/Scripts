<#
    Script Name:  Restart-DattoServices.ps1
    Author:  Rob Rathbun
    Date:  07/11/2017
    Purpose:  Simple script to fix issues with ShadowSnap services for Datto agents.
#>

$services = @("vsnapvss", "ShadowProtectSvc", "raw_agent_svc")

foreach ($service in $services) {
    Stop-Service -Name $service
    If (Get-Process -Name $service) {
        Stop-Service -Name $service -Force
    }
    Start-Service -Name $service
}