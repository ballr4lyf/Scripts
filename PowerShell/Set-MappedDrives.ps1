<#
    Script Name:  Set-MappedDrives.ps1
    Created By:  Rob Rathbun
    Created Date:  4/10/2017
    Purpose:  Set Mapped drives on logon to a NAS, but secure credentials using
              AES encryption.
#>

$NASUser = "SomeName"
$AESkey = Get-Content "C:\Path\to\AES.key"
$pw = "C:\Path\to\pwd.txt"
$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $NASUser, (Get-Content $pw | ConvertTo-SecureString -Key $AESkey)
$NAS_IP = "xxx.xxx.xxx.xxx"

If (!(Test-Path "S:\")) {
    New-PSDrive -Name "S" -PSProvider FileSystem -Root "\\$NAS_IP\Shared" -Description "Shared" -Credential $credentials -Persist
}

If (!(Test-Path "H:\")) {
    New-PSDrive -Name "H" -PSProvider FileSystem -Root "\\$NAS_IP\Users\$NASUser" -Description "Home" -Credential $credentials -Persist
}