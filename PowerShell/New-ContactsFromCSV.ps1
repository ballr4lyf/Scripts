<#
.Synopsis
   Create Exchange contacts from a CSV file.
.DESCRIPTION
   Automatically creates Exchange contacts from a CSV file.
   The CSV file must include the following header information:
   FirstName,MiddleInitial,LastName,EmailAddress

   Script Name: New-ContactsFromCSV.ps1
   Created by: Rob Rathbun
   Date Created: August 9, 2018

.EXAMPLE
   New-ContactsFromCSV -CSV C:\My\Csv\File.csv
.EXAMPLE
   New-ContactsFromCSV -CSV C:\My\Csv\File.csv -DestinationOU "OU=My Org Unit,DC=contoso,DC=com"
#>
function New-ContactsFromCSV
{
    [CmdletBinding()]
    # [Alias()]
    [OutputType([int])]
    Param
    (
        # Path to CSV file containing contact information
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$CSV,

        # Distinguished Name of the target Organizational Unit in AD.
        [Parameter(Mandatory=$false)]
        [string]
        $DestinationOU
    )

    Begin
    {
        Import-Module ActiveDirectory

        If ([adsi]::Exists("LDAP://$($OU)")) {
            If (!(Test-Path $CSV)) {
            Write-Output "Path to CSV file is invalid."
            Exit
            }
            $Contacts = Import-Csv $CSV
        } else {
            Write-Output "Organizational Unit does not Exist. Please create or change the Destination OU."
            Exit
        }
    }
    Process
    {
        Foreach ($contact in $Contacts) {
            If (($contact.MiddleInitial -eq $null) -or ($contact.MiddleInitial -eq " ") -or ($contact.MiddleInitial -eq "")) {
                $MiddleInitial = ""
            } else {
                $MiddleInitial = (Get-Culture).TextInfo.ToTitleCase($contact.MiddleInitial)
            }
            $FirstName = (Get-Culture).TextInfo.ToTitleCase($contact.Firstname)
            $LastName = (Get-Culture).TextInfo.ToTitleCase($contact.LastName)
            $Name = $FirstName + " " + $LastName
            $DisplayName = $Name + " (External Contact)"
            $contactAlias = "$($FirstName).$($LastName)"

            If ((Get-Recipient -Identity $contact.EmailAddress -ErrorAction SilentlyContinue) -eq $null) {
                New-MailContact -OrganizationalUnit $OU `
                                -FirstName $FirstName `
                                -Initials $MiddleInitial `
                                -LastName $LastName `
                                -Name $Name `
                                -Alias $contactAlias `
                                -DisplayName $DisplayName `
                                -ExternalEmailAddress $contact.EmailAddress `
                                -PrimarySMTPAddress $contact.EmailAddress
            } Else {
                Write-Output "Contact not created for `"$($contact.Firstname) $($contact.LastName)`". Email address already exists."
            }
        } 
    }
    End
    {
    }
}