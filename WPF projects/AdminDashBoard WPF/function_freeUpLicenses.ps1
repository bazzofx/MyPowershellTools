
function freeUpLicenses {
[CmdletBinding()]
Param([parameter(ValueFromRemainingArguments=$true)][String[]] $domain )

Write-Host "FREE UP LICENSES" -BackgroundColor Yellow -ForegroundColor Black
Write-Host "This command will search for users who are BLOCKED but currently have active licenses" -ForegroundColor Yellow
Write-Host ""
Write-Host ""
    if ($domain -eq $null){
        $domain = Read-Host "What is the domain you would like to free up licenses for?" } #--cls if

Write-Host "Searching $domain" -BackgroundColor Yellow -ForegroundColor Black                           
Write-Host "Searching for inactive users with active licenses, from the domain $domain" -ForegroundColor Yellow
$result = Get-MsolUser -All | Where-Object{$_.islicensed -eq $true -and $_.BlockCredential -eq $true -and $_.UserPrincipalName -like "*$domain*"} | Select UserPrincipalName, BlockCredential, IsLicensed

    if ($result.Count -gt 1){Write-Debug "There are inactive users found holding up licenses on the domain $domain"
    Write-Host "There are inactive users holding up licenses on the domain: $domain" -ForegroundColor Green
                      Write-Host $result -ForegroundColor Yellow
                      Write-Host "----------------------------------------------------------" -ForegroundColor Green}
    else{Write-Debug "There are no users found on the domain $domain holding up licenses"
    Write-Host "----------------------------------------------------------" -ForegroundColor Red
    Write-Host "No inactive users holding up licenses were found on the domain: $domain" -ForegroundColor Red}
}
