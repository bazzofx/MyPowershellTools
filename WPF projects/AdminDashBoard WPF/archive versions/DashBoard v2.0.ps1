Add-Type -AssemblyName System.Windows.Forms

##################################################################
#--------------LOG VARIABLES START -------------------------------
$userProfile = $env:USERPROFILE
$license = "FitzRoy:STANDARDWOFFPACK"

$logPath = "$userProfile\Documents\Reports\AdminDashBoard\"
$logGoodFile = "SUCCESS $(get-date -f dd-MM-yyy).log"
$logBadFile = "EXCEPTIONS $(get-date -f dd-MM-yyy).log"
$logGood = $logPath + $logGoodFile
$logBad = $logPath + $logBadFile
$checkPathlogGood = Test-Path $logGood
$checkPathlogBad = Test-Path $logBad
$testFolderPath = Test-Path $logPath

Try{
New-Item -Path $logPath -ItemType Directory
Write-Host "All log files will be saved $logPath" -foreground Yellow}
Catch{Write-Host "All log files will be saved $logPath" -ForegroundColor Yellow -BackgroundColor Black
Write-Host "" -ForegroundColor Yellow
Write-Host $logPath -ForegroundColor Yellow
}

if ($checkPathlogGood -eq $false -and $checkPathlogBad -eq $false){
$newFileGood = New-Item $logGood
$newFileBad = New-Item $logBad
}
Else{} 


#---------------------LOG VARIABLES END-------------------------------
######################################################################
#

#------------MINOR SUB FUNCTIONS-------------
function loadCSV{
    Try{
 $global:users = Import-CSV $filePath -Header "UserPrincipalName", "Login", "Password" |Select-Object -Skip 1
 Write-Host "[SUCCESS] Data files loaded successfully" -ForegroundColor Green
  }
  Catch{Write-Host "[WARNING] Failed to import data" -ForegroundColor red }
                  } #-- close function
function Start-log{
Add-Content $logGood "------------>                Starting Log File on               $(Get-Date)"
Add-Content $logBad "------------>                Starting Log File on               $(Get-Date)"
} #-- close function
function checkConnection{
    Get-MsolDomain  -ErrorAction SilentlyContinue | Out-Null
    if($?){ 
    Write-Host "" -ForegroundColor Green}
    else{ Write-Host "You will need to login before you use this application" -ForegroundColor Yellow
          Startlogin}
                        } #-- close function
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
    if($?){Write-Host "You are now connected to Microsoft Online 365" -ForegroundColor Green 
           $lbl_connected.Text="Connection to Office 365"
           $lbl_connected.ForeColor="green"
                Add-Content $logGood "Login to Microsoft Online 365 successfull"}
    Catch [System.Exception] {Write-Warning $Error[0]
                Add-Content $logBad $Error[0]
                Add-Content $logBad "Microsfot Online 365 was not able to connect"
                Add-Content $logBad "-------------------------------"}
    Catch {"Failed to connecte to server, please check if you are connected on the VPN"
                Add-Content $logBad "Failed to connecte to server, please check if you are connected on the VPN"
                Add-Content $logBad $Error[0]
                Add-Content $logBad "-------------------------------"
    }#close Catch
        }#close Else

} #-- close function
#------------RUN SUB FUNCTIONS-------------
clear
Start-Log
#---------------------------------------

function startLogin {
login365
#login-AzureAD
} #-- close function 
function resetPass365{
clear
checkConnection
        Try{    
        $count = 0
            ForEach ($row in $global:users)
                {
                $currentRowLogin = $global:users.Login[$count]
                $currentRowPassword = $global:users.Password[$count]
                $currentRowEmail = $global:users.UserPrincipalName[$count]

                 Try{
                     Set-MsOlUserPassword -UserPrincipalName $currentRowEmail -NewPassword $currentRowPassword -ForceChangePassword $True
                     Write-Host " The password for $currentRowEmail has been reset to ---> |  $currentRowPassword" -ForegroundColor Green
                     Write-Host ""
                     Write-Host "---------------------------"
                 }
                 Catch [System.Exception] {Write-Host ""}
                 Catch {   Write-Host "[ERROR] : $currentRowEmail PASSWORD FAIL TO UPDATED." -ForegroundColor Red     }
     
             $count +=1   
                } # --close ForLoop
#Update text for pop up
$lbl_confirmation.text = "Password reset for the accounts successfully"
#creates form
$window2.ShowDialog()
#cleans up
$window2.Dispose()

            }#close Try
        Catch{Write-Host "[ERROR] Please select the .csv file to use." -ForegroundColor White -BackgroundColor red}

} #-- close function
function unlockAccount {
clear
checkConnection


Add-Content $logGood "----------- UNLOCKING ACCOUNT SUCCESS LOG START --------------"
Add-Content $logBad "----------- UNLOCKING ACCOUNT EXCEPTION LOG START --------------"    
    Try{
        $count =0
        ForEach ($row in $global:users)
            {
            $currentRowLogin = $global:users.Login[$count]
            $currentRowPassword = $global:users.Password[$count]
            $currentRowEmail = $global:users.UserPrincipalName[$count]
            Write-Host $currentRowEmail
            Try{
            Set-Msoluser -UserPrincipalName $currentRowEmail -BlockCredential $false -ErrorAction SilentlyContinue
            Write-Host "[SUCCESS] - $currentRowEmail has been UNLOCKED" -ForegroundColor Green
            Write-Host ""
            Write-Host "---------------------------"

            Add-Content $logGood "[SUCCESS] - $currentRowEmail has been UNLOCKED"
            }
            Catch [System.Exception] {Write-Host ""}
            Catch {Write-Host "[ERROR] - Something went wrong while UNLOCKING $currentRowEmail" -ForegroundColor Red
            Add-Content $logBad "[ERROR] - Something went wrong while UNLOCKING $currentRowEmail"        
            }
            $count += 1

            } # -Close ForLoop
            Add-Content $logGood "----------- UNLOCKING ACCOUNT SUCCESS LOG ENDS --------------"
            Add-Content $logGood ""
            Add-Content $logBad "----------- UNLOCKING ACCOUNT EXCEPTION LOG ENDS --------------"
            Add-Content $logBad ""

        } #close Try



    Catch{Write-Host "[ERROR] Please select the .csv file to use." -ForegroundColor White -BackgroundColor red}
 
} #-- close function
function addLicense {
clear
checkConnection
        Add-Content $logGood "----------- ADDING LICENSES SUCCESS LOG START --------------"
        Add-Content $logGood ""
        Add-Content $logBad "----------- ADDING LICENSES EXCEPTION LOG START --------------"
        Add-Content $logBad ""
    Try{
        $count = 0
        ForEach ($row in $global:users)
            {
            $currentRowLogin = $global:users.Login[$count]
            $currentRowPassword = $global:users.Password[$count]
            $currentRowEmail = $global:users.UserPrincipalName[$count]
            $license = $global:license

            Try{
         Set-MSoluSer -UserPrincipalName $currentRowEmail -UsageLocation "GB"
         Set-MsolUserLicense -UserPrincipalName $currentRowEmail -AddLicenses $license 
         Write-Host "[SUCCESS] - License $license added to $currentRowEmail" -ForegroundColor Green
         Write-Host ""
         Write-Host "---------------------------"

         Add-Content $logGood "[SUCCESS] - License $license added to $currentRowEmail"
       }
        Catch [System.Exception] {Write-Host ""}
        Catch {Write-Host "[ERROR] - Something went when attemption to add '$license' license to $currentRowEmail" -ForegroundColor Red

        Add-Content $logBad "[ERROR] - Something went when attemption to add '$license' license to $currentRowEmail"
    }
        $count+=1

        } #-close forLoop

            Add-Content $logGood "----------- ADDING LICENSES SUCCESS LOG END --------------"
            Add-Content $logGood ""
            Add-Content $logBad "----------- ADDING LICENSES EXCEPTION LOG END --------------"
            Add-Content $logBad ""      
       
        }#close try
    Catch{Write-Host "[ERROR] You need to load the .csv file first." -ForegroundColor White -BackgroundColor red}
        } #-- close function
function lockAccount {
clear
checkConnection
Add-Content $logGood "----------- BLOCKING ACCOUNT SUCCESS LOG START --------------"
Add-Content $logGood ""
Add-Content $logBad "----------- BLOCKING ACCOUNT EXCEPTION LOG START --------------"
Add-Content $logBad ""
    Try{
        $count =0
        ForEach ($row in $global:users)
            {
            $currentRowLogin = $global:users.Login[$count]
            $currentRowPassword = $global:users.Password[$count]
            $currentRowEmail = $global:users.UserPrincipalName[$count]
        
            Try{
            Set-Msoluser -UserPrincipalName $currentRowEmail -BlockCredential $true
            Write-Host "[SUCCESS] - $currentRowEmail has been BLOCKED" -ForegroundColor yellow
            Write-Host ""
            Write-Host "---------------------------"

            Add-Content $logGood "[SUCCESS] - $currentRowEmail has been BLOCKED"
            }
            Catch [System.Exception] {Write-Host ""}
            Catch {Write-Host "[ERROR] - Something went while trying to BLOCK $currentRowEmail" -ForegroundColor Red
                                      Add-Content $logBad $Error[0]
                                      Add-Content $logBad "[ERROR] - Something went while trying to BLOCK $currentRowEmail"  }
            $count += 1
            }#- Close ForLoop
            Add-Content $logGood "----------- BLOCKING ACCOUNT SUCCESS LOG ENDS --------------"
            Add-Content $logGood ""
            Add-Content $logBad "----------- BLOCKING ACCOUNT EXCEPTION LOG ENDS --------------"
            Add-Content $logBad ""

        }
    Catch{Write-Host "[ERROR] Please select the .csv file to use." -ForegroundColor White -BackgroundColor red}
} #-- close function
function removeAllLicenses {
checkConnection
        Add-Content $logGood "----------- REMOVAL ALL LICENSES SUCCESS LOG START --------------"
        Add-Content $logGood ""
        Add-Content $logBad "----------- REMOVAL ALL LICENSES EXCEPTION LOG START --------------"
        Add-Content $logBad ""    
        Try{   
        $count = 0
        ForEach ($row in $global:users)
            {
            $currentRowLogin = $global:users.Login[$count]
            $currentRowPassword = $global:users.Password[$count]
            $currentRowEmail = $global:users.UserPrincipalName[$count]
            $license = "FitzRoy:STANDARDWOFFPACK"
    Try{        
        (get-MsolUser -UserPrincipalName $currentRowEmail).licenses.AccountSkuId |
        foreach{
            Set-MsolUserLicense -UserPrincipalName $currentRowEmail -RemoveLicenses $_
            Write-Host "[SUCCESS] - The Licenses from $currentRowEmail has been removed." -ForegroundColor yellow

            Add-Content $logGood "[SUCCESS] - The Licenses from $currentRowEmail has been removed."

            }
       }
            Catch [System.Exception] {Write-Host ""}
            Catch {Write-Host "[ERROR] - Something went wrong while removing licenses from $email" -ForegroundColor Red
            Add-Content $logBad $Error[0]
            Add-Content $logBad "[ERROR] - Something went wrong while removing licenses from $currentRowEmail"
    }
           $count +=1
    } # -- close ForEach

            Add-Content $logGood "----------- REMOVAL ALL LICENSES SUCCESS LOG END --------------"
            Add-Content $logGood ""
            Add-Content $logBad "----------- REMOVAL ALL LICENSES EXCEPTION LOG END --------------"
            Add-Content $logBad "" 
           
            
            }#close main Try
    Catch{Write-Host "[ERROR] Please select the .csv file to use." -ForegroundColor White -BackgroundColor red}
        } #--close function



### FORM DESIGN STARTS HERE

$FormObject = [System.Windows.Forms.Form]
$LabelObject = [System.Windows.Forms.Label]
$FileDialogObject = [System.Windows.Forms.OpenFileDialog]
$ButtonObject = [System.Windows.Forms.Button]
$RadioObject = [System.Windows.Forms.RadioButton]

$titleFont = "Cambria,11,style=bold"
$radioFont = "Verdana,9"
$fontBasic = "Verdana,10,style=bold"
###MAIN WINDOW
$window = New-Object $FormObject
$window.MaximumSize="360,490"
$window.MinimumSize="360,490"
$window.ClientSize="350,450"
$window.Text="Sniper Action"
$window.backcolor = "lightgray"
####$window.Add_FormClosing( { $_.Cancel = $true; } ) ###make user unable to close window

    #Define Main Title of Program
$lbl_title = New-Object $LabelObject
$lbl_title.text = "User Admin Dashboard v1.5"
$lbl_title.Autosize = $true
$lbl_title.Font="Cambria,12,style=bold" ### --------------------------------------------font here
$lbl_title.Location = New-Object system.drawing.point (35,5)

    #Define FilePath Label info
$lbl_filePath = New-Object $LabelObject
$lbl_filePath.text = "File not selected yet"
$lbl_filePath.AllowDrop = $true
$lbl_filePath.AutoEllipsis= $true
$lbl_filePath.AutoSize=$false
$lbl_filePath.BorderStyle="Fixed3D"
$lbl_filePath.width ="220"
$lbl_filePath.Height="20"
$lbl_filePath.Font="Verdana,8,style=bold"
$lbl_filePath.ForeColor="gray"
$lbl_filePath.Location = New-Object system.drawing.point (20,85)
    
    #Define label if connected or not
$lbl_connected = New-Object $LabelObject
$lbl_connected.text = "Not yet connected"
$lbl_connected.ForeColor="red"
$lbl_connected.Font="Verdana,9"
$lbl_connected.AutoSize=$true
$lbl_connected.Location = New-Object system.drawing.point (20,47)    
    #Define Search Button
$btn_Search = New-Object $ButtonObject
$btn_Search.text="Search..."
$btn_Search.AutoSize=$true
$btn_Search.Location = New-Object System.Drawing.Point(250,80)

    #Define Onboarding Label
$lbl_Onboarding = New-Object $LabelObject
$lbl_Onboarding.text = "_____________Onboarding_____________"
$lbl_Onboarding.Autosize = $true
$lbl_Onboarding.Font=$titleFont ### --------------------------------------------font here
$lbl_Onboarding.Location = New-Object system.drawing.point (20,120)

    #Define Radio options for Onboarding
$rdn_Button1 = New-Object $RadioObject
$rdn_Button1.Text ="Reset Password"
$rdn_Button1.Font=$radioFont
$rdn_Button1.Location= New-Object System.Drawing.Point(20,150)
$rdn_Button1.AutoSize=$true
$rdn_Button1.Parent="onboarding"

$rdn_Button2 = New-Object $RadioObject
$rdn_Button2.Text ="Unlock Account"
$rdn_Button2.Font=$radioFont
$rdn_Button2.Location= New-Object System.Drawing.Point(20,175)
$rdn_Button2.AutoSize=$true
$rdn_Button2.Parent="onboarding"

$rdn_Button3 = New-Object $RadioObject
$rdn_Button3.Text ="Add MS License"
$rdn_Button3.Font=$radioFont
$rdn_Button3.Location= New-Object System.Drawing.Point(20,200)
$rdn_Button3.AutoSize=$true
$rdn_Button3.Parent="onboarding"
    
    #Define Onboarding Label
$lbl_Offboarding = New-Object $LabelObject
$lbl_Offboarding.text = "_____________Offboarding_____________"
$lbl_Offboarding.Autosize = $true
$lbl_Offboarding.Font=$titleFont ### --------------------------------------------font here
$lbl_Offboarding.Location = New-Object system.drawing.point (20,230)

    #Define Radio options for Offboarding
$rdn_Button4 = New-Object $RadioObject
$rdn_Button4.Text ="Lock Account"
$rdn_Button4.Font=$radioFont
$rdn_Button4.Location= New-Object System.Drawing.Point(20,260)
$rdn_Button4.AutoSize=$true
$rdn_Button4.Parent="offboarding"

$rdn_Button5 = New-Object $RadioObject
$rdn_Button5.Text ="Remove All Licenses"
$rdn_Button5.Font=$radioFont
$rdn_Button5.Location= New-Object System.Drawing.Point(20,285)
$rdn_Button5.AutoSize=$true
$rdn_Button5.Parent="offboarding"

$btn_Login = New-Object $ButtonObject
$btn_Login.text="Login"
$btn_Login.AutoSize=$true
$btn_Login.Font="Arial,9,style=bold"
$btn_Login.ForeColor="green"
$btn_Login.Location = New-Object System.Drawing.Point(250,40)

$btn_exec = New-Object $ButtonObject
$btn_exec.text="Execute"
$btn_exec.Font="Arial,13,style=bold"
$btn_exec.ForeColor="green"
$btn_exec.AutoSize=$true
$btn_exec.Location = New-Object System.Drawing.Point(20,400)

$btn_cancel = New-Object $ButtonObject
$btn_cancel.text="Cancel"
$btn_cancel.Font="Arial,11,style=bold"
$btn_cancel.ForeColor="red"
$btn_cancel.AutoSize=$true
$btn_cancel.Location = New-Object System.Drawing.Point(250,400)
#Functions
function SearchDialog{
        $OpenFileDialog = New-Object $FileDialogObject
        $OpenFileDialog.InitialDirectory = $Path
        $OpenFileDialog.CheckPathExists = $true
        $OpenFileDialog.CheckFileExists = $true
        $OpenFileDialog.Title = "Please feed me"
        $OpenFileDialog.Filter = "csv files (*.csv)|*.csv" #|All files (*.*)|*.*"

        $OpenFileDialog.ShowDialog()

            if ($openFileDialog.CheckFileExists -eq $true) {
                $filePath = $openFileDialog.FileName
                $global:users = $filePath }
        $lbl_filePath.Text = $filePath
        $lbl_filePath.ForeColor="green"

loadCSV
}

###FUNCTIONS FOR FORM
function checkSelected {
if($rdn_Button1.Checked -eq $true) {resetPass365}
Elseif ($rdn_Button2.Checked  -eq $true ) {unlockAccount}
Elseif ($rdn_Button3.Checked  -eq $true ) {addLicense}
Elseif ($rdn_Button4.Checked  -eq $true ) {lockAccount}
Elseif ($rdn_Button5.Checked  -eq $true ) {removeAllLicenses}
Else {Write-Host "[ERROR] You are not yet connected to Office 365." -foreground Red} 
 }#close function

function checkConnected{
Write-Host "All log files will be saved $logPath" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------------"
Try{
    Get-MsolDomain  -ErrorAction SilentlyContinue | Out-Null
        if($?){Write-Host "You are now connected to Microsoft Online 365" -ForegroundColor Green 
               $lbl_connected.Text="Connection to Office 365"
               $lbl_connected.ForeColor="green"
                }
        Else{Write-Host "You will first need to connect to the Office 365 before using the software" -ForegroundColor Yellow}
    }#close Try
Catch{}
}#close function
checkConnected
function closeWindow{
    $window.Close()
}


##Add Functions to the Form
$btn_Search.Add_Click({SearchDialog})
$btn_Login.Add_Click({startLogin})
$btn_exec.Add_Click({checkSelected})
$btn_cancel.Add_Click({closeWindow})


##Control add things to the window
$window.Controls.AddRange(@($lbl_title,$lbl_connected,$btn_Search,$lbl_filePath,$lbl_Onboarding,$rdn_Button1,$rdn_Button2,$rdn_Button3,$lbl_Offboarding,
$rdn_Button4,$rdn_Button5,$btn_Login,$btn_exec,$btn_cancel))


#creates form
$window.ShowDialog()
#cleans up
$window.Dispose()

#checkConnected # check the connection when the programms open

