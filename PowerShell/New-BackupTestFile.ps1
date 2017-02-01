<#
    Script Name:  New-BackupTestFile.ps1
    Created By:  Rob Rathbun
    Created Date:  09/09/2016
    Purpose:  A evolution of the Create-RestoreFile.ps1 script.
              Performs the same basic job, but also creates Helpdesk
              tickets every ~90 days.

              Note (9/13/16):
              To create the AES.key file:

                  $keyFile = "C:\Path\To\AES.key"
                  $key = New-Object Byte[] 32
                  [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($key)
                  $key | Out-File $keyFile

              To create messagePW.txt

                  $messagePWFile = "C:\Path\To\messagePW.txt"
                  $key = Get-Content "C:\Path\To\AES.key"
                  $password = "MyP@ssw0rd1" | ConvertTo-SecureString -AsPlainText -Force
                  $password | ConvertFrom-SecureString -key $key | Out-File $messagePWFile
#>

############  Global Variables ############
    $drives = Get-WmiObject -Class Win32_Volume | ?{$_.DriveType -eq 3}
    $today = Get-Date
    $hostname = $env:COMPUTERNAME
    $rootFolder = "\BackupTests"
    $fileContainer = "\SourceFile"
    $fileName = "\RestoreMe.txt"
    $driveLetters = @()
    $smtpServer = "mail.mydomain.com"
    $from = "myUserName@mydomain.com"
    $to = "helpdeskSystem@mydomain.com"
    $subject = "Test backups for server: " + $hostname
###########################################

function TimeKeeperFile($driveLetter) {
    If (Test-Path ($driveLetter + $rootFolder + "\DoNotDelete")) {
        Remove-Item ($driveLetter + $rootFolder + "\DoNotDelete") -Force
    } #if
    New-Item ($driveLetter + $rootFolder + "\DoNotDelete") -ItemType File
    (Get-Item -Path ($driveLetter + $rootFolder + "\DoNotDelete")).Attributes = "Hidden","System"
} #function

function MailBodyBuilder($driveLetterArray) {
    $sources = ""
    $destinations = ""
    foreach ($Letter in $driveLetterArray) {
        $sources += "    " + $Letter + $rootFolder + $fileContainer + $fileName + "`r`n"
        $destinations += "    " + $Letter + $rootFolder + $fileName + "`r`n"
    }
    $sources += "`r`n"
    $destinations += "`r`n"
    $pre = "Recover the following file(s) from backup dated " + (Get-Date).AddDays(-29).ToString('MM/dd/yyyy') + ": `r`n`r`n"
    $mid = "To the following path(s):  `r`n`r`n"
    $post = "Document success/failure and retain documentation for 1 (one) year."
    Write-Output ($pre + $sources + $mid + $destinations + $post) | Out-String
} #function

foreach ($drive in $drives) {
    If ($drive.DriveLetter[0] -ne $null) {
        If (!(Test-Path ($drive.DriveLetter + $rootFolder + $fileContainer))) {
            New-Item -ItemType Directory -Path ($drive.DriveLetter + $rootFolder + $fileContainer)
            TimeKeeperFile($drive.DriveLetter)
        } #if

        If (Test-Path ($drive.DriveLetter + $rootFolder + $fileContainer  + $fileName)) {
            Remove-Item ($drive.DriveLetter + $rootFolder + $fileContainer  + $fileName)
        } #if

        $body = "Computer Name:  " + $hostname + "`r`nDrive Letter:  " + $drive.DriveLetter + "`r`nFile Creation Date:  " + $today
        New-Item -ItemType File -Path ($drive.DriveLetter + $rootFolder + $fileContainer  + $fileName) -Value $body

        # Check if the file has been restored.
        If (Test-Path($drive.DriveLetter + $rootFolder + $fileName)) {

            # NOTE: because this script is scheduled to run daily, I'm using a margin of error of 2 days, to allow for unforseen interuptions during script runtime.

            # If the "LastAccessTime" of the recovered file is less than (today - 2), set it to (today + 30).
            If ((Get-Item ($drive.DriveLetter + $rootFolder + $fileName).LastAccessTime) -lt $today.AddDays(-2)) {
                Set-ItemProperty -Path ($drive.DriveLetter + $rootFolder + $fileName) -Name LastAccesstime -Value ($today.AddDays(30))
            }

            # If the "LastAccessTime" of restored file is greater than today, AND less than (today + 2),
            # delete it and create new "TimeKeeperFile".
            If ((Get-Item ($drive.DriveLetter + $rootFolder + $fileName).LastAccessTime) -gt $today -and 
            (Get-Item ($drive.DriveLetter + $rootFolder + $fileName).LastAccessTime) -lt $today.AddDays(2)) {
                Remove-Item ($drive.DriveLetter + $rootFolder + $fileName) -Force
                TimeKeeperFile($drive.DriveLetter)
            }

        }

        If ((Get-Item -Path ($drive.DriveLetter + $rootFolder + "\DoNotDelete") -Force).CreationTime -lt (Get-Date).AddDays(-59)) {
            $driveLetters += $drive.DriveLetter
        } #if
    } #if
} #foreach

If ($driveLetters -ne $null) {
    $username = "someUserName"
    $key = Get-Content "C:\Path\To\AES.key"
    $messagePWFile = "C:\Path\To\messagePW.txt"
    $creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, (Get-Content $messagePWFile | ConvertTo-SecureString -Key $key)
    $body = MailBodyBuilder($driveLetters)

    Send-MailMessage -SmtpServer $smtpServer -To $to -From $from -Subject $subject -Body $body -Credential $creds
} #if