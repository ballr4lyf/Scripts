<#
.Synopsis
   Get Lockout information for specified user account.
.DESCRIPTION
   Retrieves lockout information for the referenced locked out domain account.

   Note: The domain controller MUST be configured to "Audit Account Lockout" successes and failures.
         This can be found in Group Policies under the [Computer Configuration\Windows Settings\Security
         Settings\Local Policies\Audit Policy\Audit account management] policy.
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
        # The PDC role is no longer the sole source for lockout information in a Server 2016 domain.
            # $PDC = $DCs | Where-Object{$_.OperationMasterRoles -contains "PDCEmulator"}

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

        $LockoutStats | Format-Table -Property Name,LockedOut,DomainController,BadPwdCount,AccountLockoutTime,LastBadPasswordAttempt -AutoSize

        Try
        {
            # Lockout events tracked on all DCs, not just the PDC as in previous versions of Windows Server. 
            # Also, the Event ID is now 4625 instead of 4740.
            $DCs | ForEach-Object {$LockoutEvents += Get-WinEvent -ComputerName $_.Hostname -FilterHashtable @{Logname='Security';Id=4625} `
                -ErrorAction Stop | Sort-Object -Property TimeCreated -Descending}
        }
        Catch
        {
            Write-Warning $_
            Continue
        }

        foreach ($event in $LockoutEvents)
        {
            <# Event properties have changed in Server 2016.
               Properties[5] will reference the SamAccountName of the user instead of Properties[0] referencing the SID.
               Properties[13] is where the originating machine name is located as opposed to Properties[1].
            #>
                        If ($event | Where-Object {$_.Properties[5].value -match $userStats.SamAccountName})
            {
                $event | Select-Object -Property @(
                    @{Label = 'User'; Expression = {$_.Properties[5].Value}}
                    @{Label = 'DomainController'; Expression = {$_.MachineName}}
                    @{Label = 'EventId'; Expression = {$_.Id}}
                    @{Label = 'Message'; Expression = {$_.Message -split "`r" | Select-Object -First 1}}
                    @{Label = 'LockoutLocation'; Expression = {$_.Properties[13].Value}}
                )
            }
        }

    }
    End
    {
    }
}