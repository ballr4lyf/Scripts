$oneYear = (Get-Date).AddDays(-365)
$targetFolder = "C:\Users\rrathbun\Downloads"
$Files = Get-ChildItem $targetFolder -Recurse | Where-Object {$_.LastWriteTime -le $oneYear}

foreach ($file in $Files) {
    If ($file -ne $null) {
        Write-Host "Deleting File " $file.FullName -ForegroundColor Red
        Remove-Item $file.FullName | Out-Null
    }
    else {
        Write-Host "No more files to delete" -ForegroundColor Green
    }
}