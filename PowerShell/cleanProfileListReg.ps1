
Set-Location HKLM:\

Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object {
    If ($_.Name -like "*.bak") {
        # Write-Debug "Removing: " $($_.Name)
        Remove-Item -Path $_.Name
    }
}
