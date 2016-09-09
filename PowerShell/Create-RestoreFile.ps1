<#
    Script Name:  Create-RestoreFile.ps1
    Created By:  Rob Rathbun
    Created Date:  09/01/2016
    Purpose:  Create a simple text file that can be used to test file restorations.
              File contains just a date/time entry for when it was created.
#>

$drives = Get-WmiObject -Class Win32_Volume | ?{$_.DriveType -eq 3}
$today = Get-Date
$hostname = $env:COMPUTERNAME
$driveLetters = @()

foreach ($drive in $drives) {
    If ($drive.DriveLetter[0] -ne $null) {
        $folder = $drive.DriveLetter + "\BackupTests\SourceFile"
        $file = "RestoreMe.txt"

        If (!(Test-Path $folder)) {New-Item -ItemType "directory" -Path $folder}
        If (Test-Path (Join-Path -Path $folder -ChildPath $file)) {
            Remove-Item -Path (Join-Path -Path $folder -ChildPath $file) -Force
        }
        $body = "Computer Name:  " + $hostname + "`r`nDrive Letter:  " + $drive.DriveLetter + "`r`nFile Creation Date:  " + $today

        New-Item -Path $folder -ItemType "file" -Name $file -Value $body

        If (Test-Path ($drive.DriveLetter + "\BackupTests\" + $file)) {
            Rename-Item ($drive.DriveLetter + "\BackupTests\" + $file) -NewName "Restored.txt"

            $restoredDate = (Get-Date).AddDays(-1)
            $restoredDate = Get-Date $restoredDate -Format d
            Add-Content -Path ($drive.DriveLetter + "\BackupTests\Restored.txt") -Value ("`r`n`r`nRestored:  " + $restoredDate)
        }
        If (Test-Path ($drive.DriveLetter + "\BackupTests\Restored.txt")) {
            $restoredFile = Get-Item ($drive.DriveLetter + "\BackupTests\Restored.txt")
            If ($restoredFile.LastWriteTime -gt $today.AddDays(-90)) {
                $driveLetters += $drive.DriveLetter
            }
        }
    }
}