

$env:USERNAME

$shellFolders = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

$userShellFolders = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

$oldServer = "gccba2i"

$newServer = "gccbfs01"

$profileFolders = @("Administrative Tools", "Desktop", "Favorites", "Personal", "Programs", "Start Menu", "Startup")

<# 

    get-itemproperty -path $shellFolders | ForEach-Object {
   #     Case (($_.personal).startswith($oldServer)) {
            set-itemproperty $_.personal -value $newServer\Users\$env:USERNAME
        }
        Case ($_.desktop -startswith $oldServer) {
            set-itemproperty $_.desktop -value $newServer\Users\$env:USERNAME\Desktop
        }
        Case ($_.favorites -startswith $oldServer) {
            set-itemproperty $_.favorites -value $newServer\Users\$env:USERNAME\Favorites
        }
        End Case
    }

#>