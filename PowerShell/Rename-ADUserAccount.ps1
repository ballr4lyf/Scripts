﻿<#
.Synopsis
   Rename AD Useraccount names to match a standard format based on the user's Firstname and Lastname.

   Author:  Robert Rathbun
.DESCRIPTION
   Rename AD Useraccount names to match a standard format based on the user's Firstname and Lastname.  
   
   Will also rename the user's "HomeDirectory" to match the new username.
.EXAMPLE
   Rename-ADUserAccount -userName john.doe -searchBase "CN=Org Users,DC=domain,DC=com" -format Short

   Renames useraccount "john.doe" to "jdoe".
.EXAMPLE
   Rename-ADUserAccount -userName (Get-Content C:\Users.txt) -searchBase "CN=Org Users,DC=domain,DC=com" -format Long

   Gets a list of usernames from C:\Users.txt and renames them all into the <FirstName>.<LastName> format.
#>
function Rename-ADUserAccount
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Username(s) to rename
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [string[]]$userName,

        # AD Organizational Unit to search
        [Parameter(Mandatory=$true)]
        [string]$searchBase,

        # Short or Long format of new username. Short format is <FirstInitial><LastName>.  Long format is <FirstName>.<LastName>.
        [Parameter(Mandatory=$true)]
        [ValidateSet("Short","Long")]
        [string]$format
    )

    Begin
    {
        If ($format -like "short") 
        {
            $fnLenght = 1
        } 
        Else 
        {
            $fnLenght = 12
        }
    }
    Process
    {
        $domain = Get-ADDomain
        foreach ($entry in $userName) 
        {
            try 
            {
                $user = Get-ADUser $entry -SearchBase $searchBase -Properties HomeDirectory -ErrorAction Stop
            }
            catch 
            {
                Write-Warning "Unable to get user `"" + $entry + "`". " + $($_.Exception.Message)
            }
            If ($user -ne $null) 
            {
                If ($fnLenght -eq 1) 
                {
                    $shortFirst = $user.GivenName[0]

                }
                Else
                {
                    $shortFirst = If (($user.GivenName).Length -gt 12)
                        {
                            ($user.GivenName[0..12] -join "") + "."
                        }
                        Else
                        {
                            $user.GivenName + "."
                        }
                }
                $shortLast = If (($user.Surname).Length -gt 12)
                    {
                        $user.Surname[0..12] -join ""
                    }
                    Else
                    {
                        $user.Surname
                    }
                $shortName = $shortFirst + $shortLast
                Rename-Item $user.HomeDirectory -NewName $shortName
                Set-ADUser $user -SamAccountName $shortName -UserPrincipalName ($shortName + "@" + ($domain.DNSRoot)) -HomeDirectory ($user.HomeDirectory).Replace($user.SamAccountName, $shortName)
            }
        }
    }
    End
    {
    }
}