<#
    Script Name:  Move-ComputerOU.ps1
    Created By:  Rob Rathbun
    Created Date:  07/31/2017
    Purpose:  Automatically move AD computer objects into the correct OU.
#>

$searchBase = "CN=Computers,DC=domain,DC=local"
$laptopsOU = "OU=Laptops," + $searchBase
$desktopsOU = "OU=Desktops," + $searchBase
$computers = Get-ADComputer -Filter * -SearchBase $searchBase | ?{($_.DistinguishedName -notlike $laptopsOU) -and ($_.DistinguishedName -notlike $desktopsOU)}

If ($computers -ne $null) {
    foreach ($computer in $computers) {
        If ($computer.Name -like "*WK*") {
            Move-ADObject -Identity $computer.DistinguishedName -TargetPath $desktopsOU
        } elseif ($computer.Name -like "*LT*") {
            Move-ADObject -Identity $computer.DistinguishedName -TargetPath $laptopsOU
        }
    }
}