<#
.Synopsis
   Get Lockout information for specified user account.
.DESCRIPTION
   Retrieves lockout information for the referenced locked out domain account.
.EXAMPLE
   Get-ADLockoutLocation -UserName jdoe

    Name  LockedOut DomainController  BadPwdCount AccountLockoutTime    LastBadPasswordAttempt
    ----  --------- ----------------  ----------- ------------------    ----------------------
    jdoe      True DC1.domain.com               7 2/14/2017 11:20:39 AM 2/14/2017 11:20:39 AM 
    jdoe      True DC2.domain.com               0 2/14/2017 11:20:39 AM 2/2/2017 8:42:29 AM   




    User             : jdoe
    DomainController : DC1.domain.com
    EventId          : 4740
    Message          : A user account was locked out.
    LockoutLocation  : PC3
#>
function Get-ADLockoutLocation
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Username of the locked out user account.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $UserName
    )

    Begin
    {
        $LockoutStats = @()

        Try
        {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        Catch
        {
            Write-Warning $_
            Break
        }
    }
    Process
    {
        $DCs = Get-ADDomainController -Filter *
        $PDC = $DCs | ?{$_.OperationMasterRoles -contains "PDCEmulator"}

        foreach ($DC in $DCs)
        {
            Try
            {
                $userStats = Get-ADUser -Identity $UserName -Server $DC.HostName `
                    -Properties AccountLockoutTime,LastBadPasswordAttempt,BadPwdCount,LockedOut -ErrorAction Stop
            }
            Catch
            {
                Write-Warning $_
                Continue
            }

            If ($userStats.LastBadPasswordAttempt)
            {
                $LockoutStats += New-Object -TypeName psobject -Property @{
                    Name = $userStats.SamAccountName
                    SID = $userStats.SID.Value
                    LockedOut = $userStats.LockedOut
                    BadPwdCount = $userStats.BadPwdCount
                    BadPasswordTime = $userStats.BadPasswordTime
                    DomainController = $DC.Hostname
                    AccountLockoutTime = $userStats.AccountLockoutTime
                    LastBadPasswordAttempt = ($userStats.LastBadPasswordAttempt).ToLocalTime()
                }
            }
        }

        $LockoutStats | FT -Property Name,LockedOut,DomainController,BadPwdCount,AccountLockoutTime,LastBadPasswordAttempt -AutoSize

        Try
        {
            $LockoutEvents = Get-WinEvent -ComputerName $PDC.Hostname -FilterHashtable @{Logname='Security';Id=4740} `
                -ErrorAction Stop | Sort-Object -Property TimeCreated -Descending
        }
        Catch
        {
            Write-Warning $_
            Continue
        }

        foreach ($event in $LockoutEvents)
        {
            If ($event | ?{$_.Properties[2].value -match $userStats.SID.Value})
            {
                $event | select -Property @(
                    @{Label = 'User'; Expression = {$_.Properties[0].Value}}
                    @{Label = 'DomainController'; Expression = {$_.MachineName}}
                    @{Label = 'EventId'; Expression = {$_.Id}}
                    @{Label = 'Message'; Expression = {$_.Message -split "`r" | Select -First 1}}
                    @{Label = 'LockoutLocation'; Expression = {$_.Properties[1].Value}}
                )
            }
        }

    }
    End
    {
    }
}