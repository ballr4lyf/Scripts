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

If (Get-Process -Name iexplore) {
    Stop-Process -Name iexplore -Force
}

If (!(Get-PSDrive HKU)) {
    New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS
}

Set-Location HKU:\

Get-ChildItem -Path HKU:\ | ForEach-Object {
    If (($_.Name -match '.\S-[0-9]-.') -and (($_.Name).Length -gt 30) -and !($_.Name -match '.X*Classes')) {
        If (Test-Path ($_.PSPath + "\Software\Microsoft\Internet Explorer")) {
            Remove-Item ($_.PSPath + "\Software\Microsoft\Internet Explorer") -Recurse -Force
        }
    }
}