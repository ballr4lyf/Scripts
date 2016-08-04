<#
    Script Name:  Reset-DefaultIEConfig
    Written by:  Robert D. Rathbun
    Date Created:  08/04/2016
    Purpose:  Resets Internet Explorer to its original state.
#>

If (Get-Process -Name iexplore) {
    Stop-Process -Name iexplore -Force
}

RunDll32.exe InetCpl.cpl,ResetIEtoDefaults