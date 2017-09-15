Import-Module ActiveDirectory

$Domain = (Get-ADDomain).Name
$Username = "DattoAgent"
$PW = "1qa2wsfr43ed!Q#EF$"

If (!(Get-ADUser -Identity $Username)) {
    New-ADUser -DisplayName "Datto Agent" -SamAccountName $Username -AccountPassword (ConvertTo-SecureString -AsPlainText $PW -Force)
    Add-ADGroupMember -Identity "PC Admins" -Member "$Domain\$Username"
}

$computers = @()
$services = @("stc_raw_agent", "vsnapvss", "ShadowprotectSvc")

foreach ($computer in $computers){
    foreach ($service in $services) {
        $Svc = Get-WmiObject -ComputerName $computer -Class Win32_Service -Filter "Name='$service'" 
        $Svc.StopService()
        $Svc.change($null,$null,$null,$null,$null,$null,$Domain + "\" + $Username,$PW,$null,$null,$null)
        $svc.StartService()
    }
}