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
function Verb-Noun
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Path to CSV file containing contact information
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $CSV,

        # Distinguished Name of the target Organizational Unit in AD.
        [Parameter(Mandatory=$true)]
        [string]
        $DestinationOU
    )

    Begin
    {
        Import-Module ActiveDirectory

        If ([adsi]::Exists("LDAP://$OU")) {
            If (!(Test-Path $CSV)) {
            Write-Error "Path to CSV file is invalid."
            Exit
            }
            $Contacts = Import-Csv $CSV
        } else {
            Write-Error "Organizational Unit does not Exist. Please create or change the Destination OU."
            Exit
        }
    }
    Process
    {
        Foreach ($contact in $Contacts) {
            If (($contact.MiddleInitial -eq $null) -or ($contact.MiddleInitial -eq " ") -or ($contact.MiddleInitial -eq "")) {
                $MiddleInitial = ""
            } else {
                $MiddleInitial = $contact.MiddleInitial
            }
            $Name = "$($contact.FirstName) $($contact.LastName)"
            $DisplayName = "$($contact.FirstName) $($contact.LastName) (External Contact)"
            $contactAlias = "$($contact.FirstName).$($contact.LastName)"

            If ((Get-Recipient -Identity $contact.EmailAddress) -ne $null) {
                New-MailContact -OrganizationalUnit $OU `
                                -FirstName $contact.FirstName `
                                -Initials $MiddleInitial `
                                -LastName $contact.LastName `
                                -Name $Name `
                                -Alias $contactAlias `
                                -DisplayName $DisplayName `
                                -ExternalEmailAddress $contact.EmailAddress `
                                -PrimarySMTPAddress $contact.EmailAddress
            } Else {
                Write-Output "Contact not created for `"$($contact.Firstname) $($contact.LastName)`". Please create the contact manually."
            }
        } 
    }
    End
    {
    }
}