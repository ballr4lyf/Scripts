<#
    Script Name:  Start-iisSmtpService.ps1
    Created By:  Rob Rathbun
    Created Date:  11/03/2016
    Purpose:  Script checkes the current state of the IIS SMTP service.
              Will start the service if not running.
#>

$SMTPService = [adsi]"IIS://localhost/SMTPSVC/1"

If ($SMTPService.ServerState -ne 2) {
    $SMTPService.ServerState = 2
    $SMTPService.SetInfo()
}