<#
    Script Name:  Get-ListeningPorts
    Created By:  Robert Rathbun
    Created Date:  8/23/2016
    Purpose:  Utilize PowerShell to enumerate the Listening Ports on a machine.
    
#>


$IPGlobalProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()
$TcpConnections = $IPGlobalProperties.GetActiveTcpListeners()
$UdpConnections = $IPGlobalProperties.GetActiveUdpListeners()

foreach($TcpConnection in $TcpConnections) {
    if ($TcpConnection.AddressFamily -eq "InterNetwork") {$IPType = "IPv4"} else {$IPType = "IPv6"}
    $OutputObj = New-Object -TypeName psobject
    $OutputObj | Add-Member -MemberType NoteProperty -Name "Type" -Value "TCP"
    $OutputObj | Add-Member -MemberType NoteProperty -Name "LocalAddress" -Value $TcpConnection.Address
    $OutputObj | Add-Member -MemberType NoteProperty -Name "ListeningPort" -Value $TcpConnection.Port
    $OutputObj | Add-Member -MemberType NoteProperty -Name "IPV4or6" -Value $IPType
    $OutputObj
}


<#

# This section apparently doesn't do what I expected it to do.
# It will probably removed in a future itteration.

foreach($UdpConnection in $UdpConnections) {
    if ($UdpConnection.AddressFamily -eq "InterNetwork") {$IPType = "IPv4"} else {$IPType = "IPv6"}
    $OutputObj = New-Object -TypeName psobject
    $OutputObj | Add-Member -MemberType NoteProperty -Name "Type" -Value "UDP"
    $OutputObj | Add-Member -MemberType NoteProperty -Name "LocalAddress" -Value $TcpConnection.Address
    $OutputObj | Add-Member -MemberType NoteProperty -Name "ListeningPort" -Value $TcpConnection.Port
    $OutputObj | Add-Member -MemberType NoteProperty -Name "IPV4or6" -Value $IPType
    $OutputObj
}

#>