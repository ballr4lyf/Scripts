$ComputerSystem = Get-WmiObject -Class Win32_ComputerSystem
$BIOS = Get-WmiObject -Class Win32_BIOS
$OperatingSystem = Get-WmiObject -Class Win32_OperatingSystem
$Processor =  Get-WmiObject -Class Win32_Processor
$RAMCapacity = $ComputerSystem | Measure-Object -Property TotalPhysicalMemory -Sum | %{[Math]::Round(($_.sum/1MB),2)}

Write-Host " "
Write-Host " "
Write-Host "Computer Name:     " $OperatingSystem.CSName
Write-Host "Manufacturer:      " $ComputerSystem.Manufacturer
Write-Host "Model:             " $ComputerSystem.Model
Write-Host "Domain:            " $ComputerSystem.Domain
Write-Host "Serial Number:     " $BIOS.SerialNumber
Write-Host "Operating System:  " $OperatingSystem.Caption
Write-Host "Service Pack:      " $OperatingSystem.CSDVersion
Write-Host "OS Install Date:   " ([wmi]"").ConvertToDateTime($OperatingSystem.InstallDate)
Write-Host "CPU Model:         " $Processor.Name
Write-Host "CPU Cores:         " $Processor.NumberofCores
Write-Host "CPU Threads:       " $Processor.NumberofLogicalProcessors
Write-Host "Total RAM:         " $RAMCapacity "MB"
Write-Host "Last Boot Up Time: " ([wmi]"").ConvertToDateTime($OperatingSystem.LastBootUpTime) 