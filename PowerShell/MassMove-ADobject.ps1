<#
    Requires a CSV file formated as follows:

    #Version 1.0
    Source,DestinationOU
    "CN=ComputerOrUserName1,OU=SourceOU,DC=domainname,DC=tld","OU=DestinationOU,DC=domainname,DC=tld"
    "CN=ComputerOrUserName2,OU=SourceOU,DC=domainname,DC=tld","OU=DestinationOU,DC=domainname,DC=tld"
#>

Import-Module ActiveDirectory

$myTable = Import-Csv -Path "C:\Path\to\File.csv"
$ADoHashTable = @{}

ForEach-Object ($r in $myTable) {
    $ADoHashTable[$r.Source] = $r.DestinationOU
}

ForEach-Object ($ADobject in $ADoHashTable) {
    Move-ADObject -Identity $ADobject.Source -TargetPath $ADobject.DestinationOU
}