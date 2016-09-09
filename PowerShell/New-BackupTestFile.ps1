﻿<#
    Script Name:  New-BackupTestFile.ps1
    Created By:  Rob Rathbun
    Created Date:  09/09/2016
    Purpose:  A evolution of the Create-RestoreFile.ps1 script.
              Performs the same basic job, but also creates Helpdesk
              tickets every ~90 days.
#>

############  Global Variables ############
    $drives = Get-WmiObject -Class Win32_Volume | ?{$_.DriveType -eq 3}
    $today = Get-Date
    $hostname = $env:COMPUTERNAME
    $rootFolder = "\BackupTests"
    $fileContainer = "\SourceFile"
    $fileName = "RestoreMe.txt"
    $driveLetters = @()
    $smtpServer = "mail.mydomain.com"
    $from = "myserver@mydomain.com"
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

foreach ($drive in $drives) {
    If ($drive.DriveLetter[0] -ne $null) {
        If (!(Test-Path ($drive.DriveLetter + $rootFolder + $fileContainer))) {
            New-Item -ItemType Directory -Path ($drive.DriveLetter + $rootFolder + $fileContainer)
            TimeKeeperFile($drive.DriveLetter)
        } #if

        If (Test-Path ($drive.DriveLetter + $rootFolder + $fileContainer + "\" + $fileName)) {
            Remove-Item ($drive.DriveLetter + $rootFolder + $fileContainer + "\" + $fileName)
        }

        $body = "Computer Name:  " + $hostname + "`r`nDrive Letter:  " + $drive.DriveLetter + "`r`nFile Creation Date:  " + $today
        New-Item -ItemType File -Path ($drive.DriveLetter + $rootFolder + $fileContainer + "\" + $fileName) -Value $body

        If ((Get-Item -Path ($drive.DriveLetter + $rootFolder + "\DoNotDelete")).CreationTime -lt (Get-Date).AddDays(-59)) {
            $driveLetters += $drive.DriveLetter
        } #if
    } #if
} #foreach