<# ------------------------------------------------------------------------------------

    Script Automates the start of Required Exchange Services
    
        Created By:  Robert D. Rathbun
        Company:  Networks of Florida
        Created Date:  07/22/2014
        Modified Date:  null

------------------------------------------------------------------------------------ #>

$ServiceArray = @("MSExchangeAB",
                  "MSExchangeADTopology",
                  "MSExchangeAntispamUpdate",
                  "MSExchangeEdgeSync",
                  "MSExchangeFBA",
                  "MSExchangeFDS",
                  "MSExchangeIS",
                  "MSExchangeMailboxAssistants",
                  "MSExchangeMailboxReplication",
                  "MSExchangeMailSubmission",
                  "MSExchangeProtectedServiceHost",
                  "MSExchangeRepl",
                  "MSExchangeRPC",
                  "MSExchangeSA",
                  "MSExchangeSearch",
                  "MSExchangeServiceHost",
                  "MSExchangeThrottling",
                  "MSExchangeTransport",
                  "MSExchangeTransportLogSearch",
                  "msftesql-Exchange")

foreach ($ArrMember in $ServiceArray) {
    $Service = Get-Service -Name $ArrMember
        If ($Service.Status -eq "Stopped"){
            Start-Service -Name $Service.Name
        }
}