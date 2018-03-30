<#
    Script Name: Repair-ServiceHealth.ps1
    Created By: Robert Rathbun
    Created Date: 03/30/2018

    Purpose: Check the status of important services and start them as necessary
#>

# Create an array of service names.
$services = @("ServiceName1","ServiceName2")

# Loop through each service name.
foreach ($service in $services) {
    $i = Get-Service -Name $service
    # Check if service is stopped.
    If ($i.Status -eq "Stopped") {
        # Start service if it was stopped.
        try {
            Start-Service -Name $i.Name
        } 
        catch {
            Write-Warning $_
            Continue
        }
    }
}