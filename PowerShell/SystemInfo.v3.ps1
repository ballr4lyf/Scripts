$ComputerSystem = Get-WmiObject -Class Win32_ComputerSystem
$BIOS = Get-WmiObject -Class Win32_BIOS
$OperatingSystem = Get-WmiObject -Class Win32_OperatingSystem
$Processor =  Get-WmiObject -Class Win32_Processor
$RAMCapacity = $ComputerSystem | Measure-Object -Property TotalPhysicalMemory -Sum | %{[Math]::Round(($_.sum/1MB),2)}

$ResultObject = [PSCustomObject]@{
  ComputerName = $OperatingSystem.CSName
  Manufacturer = $ComputerSystem.Manufacturer
  Model = $ComputerSystem.Model
  Domain = $ComputerSystem.Domain
  SerialNumber = $BIOS.SerialNumber
  OperatingSystem = $OperatingSystem.Caption
  ServicePack = $OperatingSystem.CSDVersion
  OSInstallDate = ([wmi]"").ConvertToDateTime($OperatingSystem.InstallDate)
  CPUModel = $Processor.Name
  CPUCores = $Processor.NumberofCores
  CPUThreads = $Processor.NumberofLogicalProcessors
  RAMinMB = $RAMCapacity
  LastStartUpTime = ([wmi]"").ConvertToDateTime($OperatingSystem.LastBootUpTime) 
}

$ResultObject
