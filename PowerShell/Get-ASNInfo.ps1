<#
Obtain ASN info of Ip address
Invoke-RestMethod -Method Get -Uri https://api.iptoasn.com/v1/as/ip/<ip Address>
Returns JSON data.

Example Return:
{
  "announced": true,
  "as_country_code": "US",
  "as_description": "LEVEL3 - Level 3 Communications, Inc.",
  "as_number": 3356,
  "first_ip": "4.0.0.0",
  "last_ip": "4.23.87.255"
}

First use "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" to set Invoke-RestMethod to use TLS 1.2.
#>