<#
    To generate the "Apps.csv" file, run the following powershell command:
        Get-AppxProvisionedPackage -Online | Select DisplayName | Export-CSV "C:\Apps.csv"

    Remove any apps from the list that you want to keep.  
    
    In this case 3dBuilder, AppConnector, OneNote, Sway, Photos, Alarms, Calculator, Camera, and WindowsCommunication Apps were retained.
#>

$appList = Import-Csv "C:\Apps.csv" | %{$_.DisplayName}

ForEach ($name in $appList) {
    Get-AppxPackage | Where-Object {$_.name -like $name} | Remove-AppxPackage
}

ForEach ($name in $appList) {
    Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $name} | Remove-AppxProvisionedPackage -Online
}