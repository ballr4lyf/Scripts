<# ----------------------------------------------------------
	Script to automate deployment of the ERP environment.
	
	Created by:  Robert D. Rathbun
	Date Created:  November 11, 2013
	For use by:  
	
		Edits:
		07/28/2016 - Sanitized script
---------------------------------------------------------- #>

# Check ComputerName.  If it is not the correct server, exit script.
	If (([Environment]::MachineName) -ne "SomeServerName") {
		Exit
	}

# Define Variables
	$Username = [Environment]::UserName
	
	$ErpRoot = "C:\SomeFolderName"
	$ErpPath = "$ErpRoot\$UserName"
		# This variable reads the contents of "_live.txt" and enumerates the last line of the file.
		$LiveFileName = (Get-Content "$ErpRoot\erpadmin\_live.txt")[-1]

# Test if the user's folder has been created.  If so, test if the user is using the most recent file.
# If both conditions are met, exit script.
	If ((Test-Path -Path $ErpPath) -ne $true) {
		New-Item $ErpPath -type directory
	} ElseIf ((Test-Path "$ErpPath\$LiveFileName") -ne $true) {
		Get-ChildItem -Path "$ErpPath" | Select -ExpandProperty FullName | Remove-Item -Force
	} Else {
		Exit
	}

# Copy file to user's ERP Directory
Copy-Item "$ErpRoot\erpadmin\$LiveFileName" -Destination $ErpPath

# If the desktop shortcut already exists, delete the shortcut.
	If ((Test-Path "C:\Users\$UserName\Desktop\ERP.lnk") -eq $true) {
		Remove-Item "C:\Users\$UserName\Desktop\ERP.lnk"
	}

# Create the desktop shortcut.
	$WshShell = New-Object -ComObject Wscript.Shell
	$ErpShortcut = $WshShell.CreateShortcut("C:\Users\$UserName\Desktop\ERP.lnk")
		$ErpShortcut.TargetPath = $ErpPath + "\" + $LiveFileName
		$ErpShortcut.WindowStyle = 1
		$ErpShortcut.Description = "ERP Lite"
		$ErpShortcut.WorkingDirectory = $ErpPath
		$ErpShortcut.Save()