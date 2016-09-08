<#
    Script Name:  Create-RestoreFile.ps1
    Created By:  Rob Rathbun
    Created Date:  09/01/2016
    Purpose:  Create a simple text file that can be used to test file restorations.
              File contains just a date/time entry for when it was created.
#>

$drives = Get-WmiObject -Class Win32_Volume | ?{$_.DriveType -eq 3}
$DateTime = Get-Date
$hostname = $env:COMPUTERNAME

foreach ($drive in $drives) {
    If ($drive.DriveLetter[0] -ne $null) {
        $folder = $drive.DriveLetter + "\BackupTests\SourceFile"
        $file = "RestoreMe.txt"

        If (!(Test-Path $folder)) {New-Item -ItemType "directory" -Path $folder}
        If (Test-Path (Join-Path -Path $folder -ChildPath $file)) {
            Remove-Item -Path (Join-Path -Path $folder -ChildPath $file) -Force
        }
        $body = "Computer Name:  " + $hostname + "`r`n Drive Letter:  " + $drive.DriveLetter[0] + "`r`n File Creation Date:  " + $DateTime

        New-Item -Path $folder -ItemType "file" -Name $file -Value $body
    }
}