<#
.Synopsis
   Obtain information on ASN of public IP address.

   Author:  Robert Rathbun
.DESCRIPTION
   Obtain information on ASN of public IP address.  

   Directly queries the RIPE database for information regarding IP addresses (https://stat.ripe.net/docs/data_api#DataCalls).
   
.EXAMPLE
  Use to evaluate a single IP address.

   C:\Windows\System32> Get-ASNInfo -IPAddress 8.8.8.8

    ASNumber : 15169
    Owner    : GOOGLE - Google LLC
    Prefix   : 8.8.8.0/24
    Country  : US
    IP       : 8.8.8.8

.EXAMPLE
  Feed data from a CSV file into the function.

  C:\Windows\System32> $csv = Import-Csv C:\MyFirewallLog.CSV
  C:\Windows\System32> $csv | Foreach{Get-ASNInfo -IPAddress $_.SourceIP} | Format-Table

  IP            Owner                           Country ASNumber Prefix
  --            -----                           ------- -------- ------
  8.8.8.8       GOOGLE - Google LLC             US      {15169}  8.8.8.0/24
  1.1.1.1       CLOUDFLARENET - Cloudflare              {13335}
#>
function Get-ASNInfo
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # IP address to inspect.
        [Parameter(Mandatory=$true)]
        [string]$IPAddress
    )

    Begin
    {
      $IPRegex = [regex] "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
      if (!($IPRegex.Match($IPAddress)).Success) {
        throw "Invalid input. Address provided is not a valid IP address."
      }
      else {
        Write-Debug "Valid IP Address Confirmed."
      }
    }
    Process
    {
      function Search-Subnet {
        <# 
        I had to search for something that would evaluate an IP address against a prefix.

        This is one of the best I found.

        Credit for this function belongs to:
        https://github.com/omniomi/PSMailTools/blob/v0.2.0/src/Private/spf/IPInRange.ps1
        #>
        param(
          # IP address to Evaluate.
          [parameter(Mandatory, Position=0)]
          [validatescript({
            ([System.Net.IPAddress]$_).AddressFamily -eq 'InterNetwork'
          })]
          [string]
          $Address,

          # Subnet to Search.
          [parameter(Mandatory, Position=1)]
          [validatescript({
            $IP = ($_ -split '/')[0]
            $Mask = ($_ -split '/')[1]

            (([System.Net.IPAddress]($IP)).AddressFamily -eq 'InterNetwork')

            if (!($Mask)) {
              throw "Missing CIDR Bit Mask."
            } elseif (!(0..32 -contains [int]$Mask)) {
              throw "Invalid CIDR Bit Mask."
            }
          })]
          [Alias('CIDR')]
          [string]
          $Subnet
        )

        # Split the subnet into Address and Mask.
        [string]$CIDRAddress = $Subnet.Split('/')[0]
        [int]$CIDRMask = $Subnet.Split('/')[1]

        # Convert all items into Int32 and calculate the full mask from the CIDR Mask.
        [int]$Base = [System.BitConverter]::ToInt32((([System.Net.IPAddress]::Parse($CIDRAddress)).GetAddressBytes()),0)
        [int]$IP   = [System.BitConverter]::ToInt32((([System.Net.IPAddress]::Parse($Address)).GetAddressBytes()),0)
        [int]$Mask = [System.Net.IPAddress]::HostToNetworkOrder(-1 -shl (32 - $CIDRMask))

        # Determine if the address is in the subnet.
        if (($Base -band $Mask) -eq ($IP -band $Mask)){
          $true
        } else {
          $false
        }
      }

      try {
        $baseURI = "https://stat.ripe.net/data/"

        $netinfo = $baseURI + "network-info/data.json?resource=" + $IPAddress
        Write-Debug "Obtaining ASN number from URI: $($netinfo)"

        $ASN = (Invoke-RestMethod -Uri $netinfo -Method Post -ErrorAction SilentlyContinue).data.asns
        
        $ASInfo = $baseURI + "as-overview/data.json?resource=AS" + $ASN
        Write-Debug "Obaining AS Holder infor from URI: $($ASInfo)"

        $ASHolder = (Invoke-RestMethod -Uri $ASInfo -Method Post -ErrorAction SilentlyContinue).data.Holder

        $GeoInfo = $baseURI + "geoloc/data.json?resource=AS" + $ASN
        Write-Debug "Obtaining Geolocation info from URI: $($GeoInfo)"

        $locations = (Invoke-RestMethod -Uri $GeoInfo -Method Post -ErrorAction SilentlyContinue).data.locations

        $mask = 0
        [string]$country = $null

        foreach ($location in $locations) {
          $prefixes = $location.prefixes
          
          foreach ($prefix in $prefixes){
            if (Search-Subnet -Address $IPAddress -Subnet $prefix) {              
              If (($prefix.Split('/')[1]) -gt $mask) {
                $mask = ($prefix.Split('/')[1])
                $country = $location.country
                $subnet = $prefix
              }
            }
          }
        }

        $Results = [PSCustomObject]@{
          IP = $IPAddress
          ASNumber = $($ASN)
          Owner = $ASHolder
          Country = $country
          Prefix = $subnet          
        }
      }
      Catch {
        $Results = $null
        Write-Error "Unable to process request."
      }
    }
    End
    {
      return $Results
    }
}