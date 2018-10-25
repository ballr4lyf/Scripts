<#
    Script Name: Get-FailedLogonAttempts.ps1
    Created By: Rob Rathbun
    Created Date: 10/25/2018

    Description: Returns all events in the security log with Event ID 4625.
#>

$events = Get-WinEvent -FilterHashtable @{Logname='Security'; Id=4625} 

$results = @()

    foreach ($event in $events) {
        If ($event -ne $null) {
            $eventInfo = [PSCustomObject] @{
                TimeCreated = $event.TimeCreated
                UserAccount = $event.Properties[5].Value
                WorkstationName = $event.Properties[18].Value
                SourceIP = $event.Properties[19].Value
            }

            $results += $eventInfo
        }
    }


$results | Export-Csv C:\Path\to\my\file\RDPAttempts.csv