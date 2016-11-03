<#
    Script Name:  Start-iisSmtpService.ps1
    Created By:  Rob Rathbun
    Created Date:  11/03/2016
    Purpose:  Script checkes the current state of the IIS SMTP service.
              Will start the service if not running.
#>

$SMTPService = [adsi]"IIS://localhost/SMTPSVC/1"

If ($SMTPService.ServerState -ne 2) {
    $LogSource = "IIS SMTP"
    Try {
        Get-EventLog -LogName Application -Source $LogSource -ErrorAction Stop
    }
    Catch {
        New-EventLog -LogName Application -Source $LogSource
    }

    Write-EventLog -LogName Application -Source $LogSource -EntryType Warning -EventId 1 -Message "SMTP Service has stopped. Starting Service."
    $SMTPService.ServerState = 2
    $SMTPService.SetInfo()
}