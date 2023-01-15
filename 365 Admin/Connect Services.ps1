<#
set-calendarProcessing $resource`
    -AllBookInPolicy $false
    #>
#Connect Exchange Service
###Connect-ExchangeOnline -UserPrincipalName pbadmin@efitzroy.org.uk

$user = "CoventryMeetingRoom@fitzroy.org" 

Write-Host "Email in scope is: $user" -ForegroundColor Yellow 
get-calendarProcessing $user |fl
Write-Host "Email in scope is: $user" -ForegroundColor Yellow