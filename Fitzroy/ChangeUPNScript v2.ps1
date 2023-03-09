#subFunctions - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - - 
function login365{
#Login to Office 365
Get-MsolDomain -ErrorAction SilentlyContinue 
    if($?){Write-Host "You are now connected to Microsoft Online 365" -ForegroundColor Green 
           $lbl_connected.Text="Connection to Office 365"
           $lbl_connected.ForeColor="green"
            }
    Else{

    Try{
    Connect-MsolService -ErrorAction SilentlyContinue
                Get-MsolDomain -ErrorAction SilentlyContinue 
    if($?){Write-Host "You are now connected to Microsoft Online 365" -ForegroundColor Green } #close if
    } #close Try
    Catch {Write-Host "Failed to connecte to server, please check if you are connected on the VPN" -ForegroundColor Red} #close Catch

        }#close Else 
} #-- close function
function loginAD {
Write-Host "We need to login into AD also"
$ADConnection = New-PSSession azdc01 -Credential(Get-Credential)
Enter-PSSession $ADConnection

}
function checkChange{
 #check Changes occured on AD
 $EmailAddress = Get-ADUser -Filter {userPrincipalName -eq $name}  -Properties * | Select EmailAddress
 Write-Host "Email on AD has been changed to $EmailAddress" -ForegroundColor Green
}
#subFunctions - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - -  - - - 
function checkConnection{
Try{
    Get-MsolDomain  -ErrorAction SilentlyContinue | Out-Null
    if($?){ #$? means previous line of work that was successfully run, in this case if login was successful dont ask to login to 365.
    Write-Host "You are connected " -ForegroundColor Green}
    else{ Write-Host "You will need to login before you use the script" -ForegroundColor Yellow
          login365
          $token365 = $true} #close else
    }
Catch{Write-Host "You need to login into the Msol-Service first" -ForegroundColor Yellow}

Try{
    Get-ADDmain | Out-Null
    if($?){ #$? means previous line of work that was successfully run, in this case if login was successful dont ask to login to AD.
    Write-Host "You are connected " -ForegroundColor Green}
    else{ Write-Host "You will need to login to AD before you use the script" -ForegroundColor Yellow
          loginAD
          $tokenAD = $true} #close else
   }
Catch{Write-Host "You need to login into the Azdc01 first" -ForegroundColor Yellow}      

          }# --close function
function Main {
checkConnection

if($tokenAD -eq $true -and $token365 -eq $true){


#Store the data from csv in the $ADUsers variable
$ADUsers = Import-csv C:\test\ChangeUPN.csv


#               CHANGE THE UPN ON THE MICROSFOT 365 SERVER
#Loop through each row containing user details in the CSV file 
foreach ($User in $ADUsers)
{
	#Read user data from each field in each row and assign the data to a variable as below
		
	$UserPrincipalName 	= $User.UserPrincipalName
	$NewUserPrincipalName 	= $User.NewUserPrincipalName

		Set-MsolUserPrincipalName `
            -UserPrincipalName $UserPrincipalName `
            -NewUserPrincipalName $NewUserPrincipalName `

#               CHANGE THE EmailAddress Property field on AD        	
$account = Get-ADUser -Filter {userPrincipalName -eq $UserPrincipalName}  -Properties * | Select SamAccountName | Out-String
$chop = $account.Split("-") # split the word remove unecessary bits
$account = $chop[-1].trim() #get the last word from the split / trim removes white spaces
Set-ADUser -identity $account -EmailAddress $NewUserPrincipalName
checkChange
}

}## -close IF check if $tokenAD and $token365 are true

}


Main