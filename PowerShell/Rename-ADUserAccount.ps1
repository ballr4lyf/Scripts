<#
.Synopsis
   Rename AD Useraccount names to match a standard format based on the user's Firstname and Lastname.

   Author:  Robert Rathbun
.DESCRIPTION
   Rename AD Useraccount names to match a standard format based on the user's Firstname and Lastname.  
   
   Will also rename the user's "HomeDirectory" to match the new username.
.EXAMPLE
   Rename-ADUserAccount -userName john.doe -format Short

   Renames useraccount "john.doe" to "jdoe".
.EXAMPLE
   Rename-ADUserAccount -searchBase "CN=Org Users,DC=domain,DC=com" -format Short

   Rename all user accounts found under the "Org Users" OU using the <FirstInitial><LastName> format.
.EXAMPLE
   Rename-ADUserAccount -userName (Get-Content C:\Users.txt) -format Long

   Gets a list of usernames from C:\Users.txt and renames them all into the <FirstName>.<LastName> format.
#>
function Rename-ADUserAccount
{
    [CmdletBinding(DefaultParameterSetName = 'SearchBy')]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Username(s) to rename
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName = 'SearchBy')]
        [string[]]$userName,

        # AD Organizational Unit to search
        [Parameter(Mandatory=$false,
                   ParameterSetName = 'SearchBy')]
        [string]$searchBase,

        # Short or Long format of new username. Short format is <FirstInitial><LastName>.  Long format is <FirstName>.<LastName>.
        [Parameter(Mandatory=$true)]
        [ValidateSet("Short","Long")]
        [string]$format
    )

    Begin
    {
        $lastNameLenghth = 12  # set the max number of characters for the LastName portion of the username.
        If ($format -like "short") # If "short" format is chosen.
        {
            $nameLength = 1 # '$nameLength' set to 1 to only grab the first initial.
        } 
        Else # Else "long" format has been chosen.
        {
            $nameLength = $lastNameLenghth # Match length to '$lastNameLenghth'.
        }
        Import-Module ActiveDirectory

        If ($userName -eq $null) # If '-searchBase' option is used.
        {
            # Create and add users to the '$users' array.
            Try 
            {
                $users = Get-ADUser -Filter * -SearchBase $searchBase -Properties HomeDirectory -ErrorAction Stop 
            }
            Catch
            {
                Write-Warning "Unable to get userlist with Searchbase `"$searchBase`"."
            }
        }
        Else # The '-userName' option was used.
        {
            $users = @()
            foreach ($user in $userName) 
            {
                # Add users to the '$users' array.
                Try
                {
                    $users += Get-ADUser $user -Properties HomeDirectory -ErrorAction Stop 
                }
                Catch
                {
                    Write-Warning "Unable to get user with username `"$user`"." -WarningAction Continue
                }
            }
        }
    }
    Process
    {
        $domain = Get-ADDomain
        If ($users -ne $null) # check if the '$users' array is empty.
        {
            foreach ($user in $users) # iterate through each username in the '$users' array.
            {
            
                If ($user -ne $null) # check to make sure the username is not <blank>.
                {
                    $shortFirst = $user.GivenName[0..($nameLength - 1)] -join "" # set the first part of the new username.
                    If ($nameLength -ne 1) # If 'Long' format was chosen, add a "." to the end of the first part of the new username.
                    {
                        $shortFirst += "."
                    }
                    # Set the second part of the new username using the '$_.Surname' property.
                    $shortLast = If (($user.Surname).Length -gt $lastNameLenghth) # Use the defined '$lastNameLenght' variable.
                        {
                            $user.Surname[0..($lastNameLenghth - 1)] -join ""
                        }
                        Else
                        {
                            $user.Surname
                        }
                    $shortName = $shortFirst + $shortLast # Define the new username by combining the first two parts.

                    If ($shortName -notlike $user.SamAccountName) # Check to make sure the username is not already defined correctly.
                    {
                        Rename-Item $user.HomeDirectory -NewName $shortName  # Rename the HomeDirectory folder.
                        # Define arguments to be used in setting the username(s).
                        $arguments = @{SamAccountName = $shortName; 
                                       UserPrincipalName = ($shortName + "@" + ($domain.DNSRoot)); 
                                       HomeDirectory = ($user.HomeDirectory).Replace($user.SamAccountName, $shortName)}
                        Set-ADUser $user @arguments # Set username(s) to the generated username.
                    }
                    Else
                    {
                        Write-Warning "No changes made. Username already set correctly." # username was already set to the chosen 'Short' or 'Long' format.
                    }
                }
            }
        }
        Else
        {
            Write-Error "The provided parameters returned no users.  Unable to continue."
        }
    }
    End
    {
    }
}