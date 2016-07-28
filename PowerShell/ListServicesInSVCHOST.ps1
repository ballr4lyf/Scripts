<#
    ScriptName:  ListServicesInSVCHOST.ps1
    Purpose:  To be able to determine which services are running in a given "SVCHOST" process.
    Created By:  Robert D. Rathbun
    Created Date:  05/04/2016
#>

# Find all SVCHOST processes
$svchosts = Get-Process -Name svchost | select Name,Id
$services = @()

# Loop through each SVCHOST and gather the list into an array of $services
foreach($svchost in $svchosts) {
    $service = Get-WmiObject -Class Win32_Service -Filter "ProcessID='$($svchost.Id)'" | select ProcessID,Name,DisplayName,State
    $services += $service
}

# Write the array of services out to the console
Write-Output $services | FT ProcessID,Name,DisplayName,State -AutoSize