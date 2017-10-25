<#
    Requires a CSV file formated as follows:

    #Version 1.0
    Source,DestinationOU
    "CN=ComputerOrUserName1,OU=SourceOU,DC=domainname,DC=tld","OU=DestinationOU,DC=domainname,DC=tld"
    "CN=ComputerOrUserName2,OU=SourceOU,DC=domainname,DC=tld","OU=DestinationOU,DC=domainname,DC=tld"
#>

Import-Module ActiveDirectory

$content = Import-Csv -Path "C:\Path\To\File.csv"
$objectArray = @()

foreach ($item in $content) {
    $itemDetails = [System.Management.Automation.PSCustomObject]@{
        Source = $item.Source
        Destination = $item.DestinationOU
    }

    $objectArray += $itemDetails
}

ForEach-Object ($ADobject in $objectArray) {
    Move-ADObject -Identity $ADobject.Source -TargetPath $ADobject.Destination
}