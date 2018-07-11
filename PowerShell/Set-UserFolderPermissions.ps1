<#
    Script Name: Set-UserFolderPermissions.ps1
    Created By: Rob Rathbun
    Created Date: July 11, 2018
    Purpose: Recurse through the defined user folder and apply proper permissions to folders.
             Script also disables ACL inheritance on folders in order to apply permissions.
             This script assumes that the folder name matches the user's "SamAccountName" property.
#>

Import-Module ActiveDirectory

$UserFolders = Get-ChildItem -Path D:\Users
$domainNBName = $((Get-ADDomain).NetBIOSName)
$ReadGroup = "$domainNBName\Folder - Users - Read"
$ModifyGroup = "$domainNBName\Folder - Users - Modify"

foreach ($folder in $UserFolders) {
    try {
        $acl = $folder | Get-Acl

        <#Remove Inheritance. 
        Reference: https://msdn.microsoft.com/en-us/library/system.security.accesscontrol.objectsecurity.setaccessruleprotection(v=vs.110).aspx
        #>
        $acl.SetAccessRuleProtection($true,$true)
        Set-Acl -Path $($folder.FullName) -AclObject $acl
    
        #Retrieve ACL list again after disabling inheritance.    
        $acl = $folder | Get-Acl

        #Remove access from "Domain Users"
        $acl.Access | Where {$_.IdentityReference -eq "$domainNBName\Domain Users"} | Foreach{$acl.RemoveAccessRule($_)} | Out-Null

        #Check if entry exists for assigned user. If not, add "Modify" entry.
        $userACE = $acl.Access | Where{$_.IdentityReference -eq "$domainNBName\$($folder.Name)"}
        if ($userACE -eq $null) {
            $userRule = New-Object System.Security.AccessControl.FileSystemAccessRule -ArgumentList ("$domainNBName\$($folder.Name)", "Modify", "Allow")
            $acl.SetAccessRule($userRule)
        }

        #Add Rules for Groups
        $readRule = New-Object System.Security.AccessControl.FileSystemAccessRule -ArgumentList @($ReadGroup, "ReadAndExecute", "Allow")
        $modifyRule = New-Object System.Security.AccessControl.FileSystemAccessRule -ArgumentList @($ModifyGroup, "Modify", "Allow")
        $acl.SetAccessRule($readRule)
        $acl.SetAccessRule($modifyRule)

        Set-Acl -Path $($folder.FullName) -AclObject $acl
    }
    Catch {
        Write-Output "Unable to process folder: `"$($folder.FullName)`""
    }
}