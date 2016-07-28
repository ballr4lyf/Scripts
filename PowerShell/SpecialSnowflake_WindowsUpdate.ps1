<#
    SPECIALSNOWFLAKESERVER (SSS) Custom Windows Update Script

    Author:  Robert D. Rathbun
    Created Date:  2/1/2016
    Information: 
        The service named "yourServiceNameHere" on SSS will cause database corruption in the 
        yourServiceNameHere Database if the service is not shut down properly.  The most likely
        cause of an improper stop of the service is if it is downloading reports while the "shutdown"
        command is sent to the SSS.  The "shutdown" command will try to wait for the service to stop,
        but it will eventually time out and proceed with the reboot anyways.  This is what causes database
        corruption.

        This script is set download and install updates as a scheduled task, then shut down the "yourServiceNameHere"
        service, and then reboot the SSS.  Some lines are borrowed from Gregory Strike's script
        (available here:  https://community.spiceworks.com/scripts/show/1075-download-and-install-updates-with-or-without-reboot).
#>

$ServiceName = "yourServiceNameHere"

$UpdateSession = New-Object -ComObject "Microsoft.Update.Session"
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

$SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software'")

If ($SearchResult.Updates.Count -eq 0) {Exit}

$UpdatesToDownload = New-Object -ComObject "Microsoft.Update.UpdateColl"

For ($x = 0; $x -lt $SearchResult.Updates.Count; $x++) {
    $Update = $SearchResult.Updates.Item($x)
    $Null = $UpdatesToDownload.Add($Update)
}

$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$Null = $Downloader.Download()

$UpdatesToInstall = New-Object -ComObject "Microsoft.Update.UpdateColl"

For ($x = 0; $x -lt $SearchResult.Updates.Count; $x++) {
    $Update = $SearchResult.Updates.Item($x)
    $Null = $UpdatesToInstall.Add($Update)
}

$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToInstall

$InstallationResults = $Installer.Install()

If ($InstallationResults.RebootRequire -eq $true) {
    Stop-Service -Name $ServiceName
    $svc = Get-Service -Name $ServiceName
    $svc.WaitForStatus('Stopped')
    Restart-Computer -Force
}