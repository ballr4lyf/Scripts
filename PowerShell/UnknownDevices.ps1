
# Create a collection of unknown devices.
$Unknowns = Get-WmiObject -Class Win32_PNPEntity | Where-Object {$_.ConfigManagerErrorCode -ne 0 -or $_.Name -match "Standard VGA"} | select Name,DeviceID
$matched = @()

# For Each unknown device...
foreach($device in $Unknowns) {
    # Pull the Vendor ID and Device ID
    $vendorID = ($device.DeviceID | Select-String -Pattern 'VEN_....' | select -expand Matches | select -expand Value) -replace 'VEN_',''
    $deviceID = ($device.DeviceID | Select-String -Pattern 'DEV_....' | select -expand Matches | select -expand Value) -replace 'DEV_',''

    # Check for null device.
    If($deviceID.Length -eq 0) {
        Write-Verbose "Found a null device. Skipping." 
        Continue
    }

    # Search the PCI Database website for the DeviceID.
    $url = "http://www.pcidatabase.com/search.php?device_search_str=$deviceID&device_search=Search"
    Try {$results = Invoke-WebRequest $url}
    Catch [System.NotSupportedException]{Write-warning "You need to launch Internet Explorer once before running this";return}

    $matches = ($results.ParsedHtml.getElementsByTagName('p') | select -expand innerHtml).Split()[1]
    Write-Verbose "Found $matches matches"

    $htmlCells = $results.ParsedHtml.getElementsByTagName('tr')  | select -skip 4 -Property *html*
    Write-Debug "test `$htmlCells for the right values $htmlCells"

    # Find devices that also match VendorID.
    $matchinDev = ($htmlCells.InnerHtml | Select-String -Pattern $vendorID | select -expand Line).ToString().Split("`n")

    if ($matchingDev.count -ge 1){
        $matchedObject = New-Object PSObject -Property @{VendorID=$vendorID;DeviceID=$deviceID;DevMgrName=$device.Name;LikelyName=$matchingDev[1] -replace '<TD>','' -replace '</TD>',''}
        $matched += $matchedObject
    }
    else{CONTINUE}
}

If($matched -ne '') {
    Write-Output $matched | FT VendorID,DeviceID,DevMgrName,LikelyName -AutoSize
}