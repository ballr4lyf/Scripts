<#
    ScriptName:  Add-MerakiClientVPN.ps1
    Created By:  Robert Rathbun
    Created Date:  12/26/2017
    Purpose:  Deploy Meraki Client VPN.  Requires Windows 8.1 or above.
#>

#Fill in the Client Details between the quotation marks below:
$VPNName = ""    # The name of the VPN Connection that will be added.
$VPNServer = ""  # The server address of the VPN endpoint.
$PSK = ""        # The pre-shared key for the VPN connection.

# Do not edit below this line.
Add-VpnConnection -Name $VPNName `
                  -ServerAddress $VPNServer `
                  -TunnelType L2tp `
                  -L2tpPsk $PSK `
                  -EncryptionLevel Optional `
                  -AuthenticationMethod Pap `
                  -AllUserConnection `
                  -Force