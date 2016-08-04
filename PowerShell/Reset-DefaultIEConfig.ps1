<#
    Script Name:  Reset-DefaultIEConfig
    Written by:  Robert D. Rathbun
    Date Created:  08/04/2016
    Purpose:  Resets Internet Explorer to its original state.
#>

<# If (Get-Process -Name iexplore) {
    Stop-Process -Name iexplore -Force
}

RunDll32.exe InetCpl.cpl,ResetIEtoDefaults #>

If (!(Get-PSDrive -Name HKU)) {
    New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS
}

Set-Location HKU:\

foreach ($user in (Get-ChildItem -Path HKU:\)) {
    If (Test-Path ("Registry::" + $user.name + "Software\Microsoft\Internet Explorer")) {
        Remove-Item ("Registry::" + $user.name + "Software\Microsoft\Internet Explorer") -Force
    }
}