


function Rename-ADUserAccount {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SearchBase,

        [Parameter(Mandatory=$true)]
        [string[]]$userName,

        [Parameter(Mandatory=$true, ValidateSet("Short","Long"))]
        [string]$format
    )
}
