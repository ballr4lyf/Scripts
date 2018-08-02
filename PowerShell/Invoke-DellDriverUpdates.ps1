<#
.Synopsis
   Invokes "Dell Command | Update" utility to update drivers on PCs.
.DESCRIPTION
   Invokes "Dell Command | Update" utility to update drives on PCs. 
   Will not run on server Operating Systems.

   Created by: Rob Rathbun
   Created Date: August 2, 2018
.EXAMPLE
   Invoke-DellDriverUpdates

   Will detect and apply the driver updates on Dell systems.
.EXAMPLE
   Invoke-DellDriverUpdates -ForceReboot

   Will detect and apply the driver updates on Dell systems. 
   Upon completion, will force the system to reboot of a driver update requires it.
#>
function Invoke-DellDriverUpdates
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Unless $true, will not reboot the target system if required.
        [Parameter(Mandatory=$false)]
        [switch]$ForceReboot

    )
    Begin
    {
        $InstallLocation = "C:\Program Files (x86)\Dell\CommandUpdate"
        $DCU = "dcu-cli.exe"
        $DownloadFile = "Dell-Command-Update_DDVDP_WIN_2.4.0_A00.EXE"
        $DownloadDir = "C:\Temp\DCU"
        $DownloadURL = "https://downloads.dell.com/FOLDER05055451M/1/$DownloadFile"
        
        Write-Debug "Checking operating system."
        If ($((Get-WmiObject -Class Win32_OperatingSystem).Caption) -like "*Server*") {
            Write-Debug "Server operating system detected. Exiting script."
            Exit
        }

        If(!(Test-Path "$InstallLocation\$DCU")) {
            Write-Debug "DCU Not installed. Invoking installation process."
            If (!(Test-Path $DownloadDir)) {
                Write-Debug "Unable to locate `"C:\Temp\DCU`" directory. Creating directory."
                New-Item $DownloadDir -ItemType Directory -Force
            }

        If (!(Test-Path "$DownloadDir\$DownloadFile")) {
            Write-Debug "Cleaning/Removing all `"*.exe`" files from `"C:\Temp\DCU`"."
            Remove-Item "$DownloadDir\*.exe" -Force

            Write-Debug "Installer not downloaded. Downloading."
            $download = {Invoke-WebRequest -Uri $args[0] -OutFile $args[1] -TimeoutSec 600}
            Start-Job -Name Webreq -ScriptBlock $download -ArgumentList $DownloadURL,"$DownloadDir\$DownloadFile"
            Wait-Job -Name Webreq
        }

        try {
            Write-Debug "Extracting DCU installer from downloaded file."
            Start-Process -FilePath "$DownloadDir\$DownloadFile" -ArgumentList "/s /e=$DownloadDir" -Wait
        }
        catch {
            Write-Debug "Unable to extract DCU installer from downloaded file. Exiting script."
            Exit
        }
        

        try {
            $ExtractedFile = Get-ChildItem $DownloadDir | Where-Object{$_.Name -like "DCU*.exe"}
        }
        catch {
            Write-Output "Installer file not located. Exiting script."
            Exit
        }
        
        try {
            Write-Debug "Silently installing Dell Command Update."
            Start-Process -FilePath $($ExtractedFile.FullName) -ArgumentList "/S /v /qn" -Wait
        }
        catch {
            Write-Output "Unable to install Dell Command Update. Exiting Script."
            Exit
        }

        Write-Debug "Installation complete. Removing executables from `"C:\Temp\DCU`" folder."
        Remove-Item "$DownloadDir\*.exe" -Force
        }
    }
    Process
    {
        If (!(Test-Path "$InstallLocation\$DCU")) {
            Write-Debug "Product not installed. Exiting script."
            Exit
        }

        Write-Debug "Setting argument list for `"Dell Command Update`" commands."
        $ArgList = "/silent"
        If ($ForceReboot) {
            $ArgList += " /reboot"
            Write-Warning "WARNING: If required, this system will automatically reboot upon completion. `r`n
                           There will be no warning given to logged on users."
        }
        $ArgList += " /log $DownloadDir"

        Write-Debug "Attempt to run the `"Dell Command Update`" utility."
        try {
            Start-Process -FilePath "$InstallLocation\$DCU" -ArgumentList $ArgList -Wait -ErrorVariable finishingState
            If ($finishingState -ne $null) {
                $OutMessage = "Execution completed with the following error: `n " + $($finishingState)
            }
            else {
                $OutMessage = "Execution of driver updates completed successfully."
            }
            
        }
        catch {
            $OutMessage = "Failed to execute update request."
        }
    }
    End
    {
        Write-Output $OutMessage
    }
}