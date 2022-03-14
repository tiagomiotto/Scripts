$clientRequestId = [GUID]::NewGuid().Guid
$version = Invoke-RestMethod -Uri ($ws + "/version/Worldwide?clientRequestId=" + $clientRequestId)
$ws = "https://endpoints.office.com"
Invoke-RestMethod -Uri ($ws + "/changes/Worldwide/"+ $version.latest + "?clientRequestId=" + $clientRequestId)

$ws + "/changes/Worldwide/"+ $version.latest + "?clientRequestId=" + $clientRequestId

$endpointSets=Invoke-RestMethod -uri $LastURL

$datapath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path + "\O365IP_change_" + $version.latest + ".txt" 

[System.Collections.ArrayList]$urlsToRemove =  @()
[System.Collections.ArrayList]$urlsToAdd =  @()

[System.Collections.ArrayList]$urlsThatJustChangedCategory = @()


 $endpointSets | Sort-Object -Property  id| ForEach-Object {

    $urlsToRemove +=$_.remove.urls
    $urlsToAdd +=  $_.add.urls
    # find matches, add to an aray, then remove from each
}



 $urlsToAdd |  ForEach-Object {

    if($urlsToRemove.Contains($_)) {
    $urlsThatJustChangedCategory += $_
    $urlsToRemove.Remove($_)
    
    }
}

 $urlsThatJustChangedCategory |  ForEach-Object {
   $urlsToAdd.Remove($_)
}


 Write-Host "URLs to Add: "
$urlsToAdd

Write-Host "--------------"
Write-Host "--------------"

  Write-Host "URLs to Remove: "
$urlsToRemove 

Write-Host "--------------"
Write-Host "--------------"

Write-Host "URLs that changed category: "

$urlsThatJustChangedCategory

  # write output to data file
    Write-Output "Office 365 IP and UL Web Service data" | Out-File $datapath
    Write-Output "Worldwide instance" | Out-File $datapath -Append
    Write-Output "" | Out-File $datapath -Append
    Write-Output ("Version: " + $version.latest) | Out-File $datapath -Append
    Write-Output "" | Out-File $datapath -Append
    Write-Output "URLs for Proxy Server" | Out-File $datapath -Append
    Write-Output "" | Out-File $datapath -Append
    Write-Output "URLs to Add" | Out-File $datapath -Append
    Write-Output "---------------" | Out-File $datapath -Append
    $urlsToAdd | Out-File $datapath -Append

    Write-Output "---------------" | Out-File $datapath -Append
    Write-Output "" | Out-File $datapath -Append
    Write-Output "URLs to Remove" | Out-File $datapath -Append
    Write-Output "---------------" | Out-File $datapath -Append
    $urlsToRemove | Out-File $datapath -Append
       Write-Output "---------------" | Out-File $datapath -Append
    Write-Output "" | Out-File $datapath -Append
    Write-Output "URLs that just changed category to be evaluated for optimization" | Out-File $datapath -Append
    Write-Output "" | Out-File $datapath -Append
    $urlsThatJustChangedCategory | Out-File $datapath -Append



Write-Host "Changes file created at: " $datapath


