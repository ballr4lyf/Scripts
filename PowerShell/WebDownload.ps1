$url = "http://mirror.internode.on.net/pub/test/10meg.test"
$InvokeOutPut = "$PSScriptRoot\invoke_10meg.test"
$WebClientOutput = "$PSScriptRoot\wc_10meg.test"
$BITSOutput = "$PSScriptRoot\bits_10meg.test"

$InvokeTime = Get-Date
Invoke-WebRequest -Uri $url -OutFile $InvokeOutPut
Write-Output "Invoke-WebRequest Time Taken: $((Get-Date).Subtract($InvokeTime).Seconds) second(s)"
#Result:  7 seconds

$webClientTime = Get-Date
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($url, $WebClientOutput)
Write-Output "WC Time Taken: $((Get-Date).Subtract($webClientTime).Seconds) second(s)"
#Result:  2 seconds

$BITSTime = Get-Date
Start-BitsTransfer -Source $url -Destination $BITSOutput
Write-Output "BITS Time Taken: $((Get-Date).Subtract($BITSTime).Seconds) second(s)"
#Result:  6 seconds 

<#
    NOTE(s):  
    1. Although the BITS transfer was slower, there is more flexibility in using the method, such as being able to retry on failure,
       and the ability to pass credentials for a web proxy AND the destination server, etc.  The downside is that it must be enabled,
       and jobs can be queued in the background, hindering execution of the script.
    2. There is no visible progress indicator for using the WebClient object.
    3. Invoke-WebRequest relies on IE, therefore cannot be used on Windows Server Core.
#>