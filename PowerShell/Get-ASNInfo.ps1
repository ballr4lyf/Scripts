<#
.Synopsis
   Obtain information on ASN of public IP address.

   Author:  Robert Rathbun
.DESCRIPTION
   Obtain information on ASN of public IP address.  

   Future updates will directly query the RIPE database (https://stat.ripe.net/docs/data_api#AsOverview).
   
.EXAMPLE
   Get-ASNInfo -IPAddress 8.8.8.8

    announced       : True
    as_country_code : US
    as_description  : GOOGLE - Google LLC
    as_number       : 15169
    first_ip        : 8.8.8.0
    ip              : 8.8.8.8
    last_ip         : 8.8.8.255   
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
        Write-Debug "IP Address Confirmed."
      }
    }
    Process
    {
      [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
      $uri = "https://api.iptoasn.com/v1/as/ip/"
      $uri += $IPAddress

      Write-Debug "Complete URI is $($uri)"

      $Results = Invoke-RestMethod -Method Get -Uri $uri -ErrorVariable RestError
      If ($RestError -ne $null) {
        return $RestError
      }
    }
    End
    {
      return $Results
    }
}