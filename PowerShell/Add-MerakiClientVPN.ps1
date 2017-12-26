<#
    ScriptName:  Add-MerakiClientVPN.ps1
    Created By:  Robert Rathbun
    Created Date:  12/26/2017
    Purpose:  Deploy Meraki Client VPN.  Requires Powershell Version 5.
#>


$VPNName = 
$VPNServer = 
$PSK = 

Add-VpnConnection -Name $VPNName `
                  -ServerAddress $VPNServer `
                  -TunnelType L2tp `
                  -L2tpPsk $PSK `
                  -EncryptionLevel Optional `
                  -AuthenticationMethod Pap `
                  -AllUserConnection `
                  -Force