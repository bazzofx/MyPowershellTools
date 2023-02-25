#------------------------------------------------------
$yesterday = (get-date).AddDays(-1); $x =[string]$yesterday; $x = $x.Substring(0,10); $yesterday = $x.replace("`/" ,"-")
$art = @"
 
       _,="(  // )"=,_
    _,'    \_>'\_/    ',_ 
     .7,     {   }     ,\.
      '/:,  .m  m.  ,:\'
        ')",(/   \),"('
            '{'!!'}' - All is looking good here, you should treat yourself with a coffee         
           (      )       ) 
            )    (        (         The report is empty. ~ No Starters/Leavers  yesterday
        .-'------------------|                                                  $yesterday
       ( C|/\/\/\/\/\/\/|
        '-./\/\/\/\/\/\/\/|  
          '_____________'                                
           '------------'
"@ #Drawing

$today = Get-Date; $x =[string]$today; $x = $x.Substring(0,10); $today = $x.replace("`/" ,"-")
$wd = "C:\iSAPScripts\DataTrackerPhase1"
$todayFilePath = "$wd\$today.csv"
$SFTPGetFilesCMD = "$wd\getfiles3.sftp" #SFTP commands  this is a .txt file save with diferent extension
#----------------------------------------------------
Function SFTPGetFiles
{
    
    Try
    {
        set-location -path "C:\Program Files (x86)\WinSCP"
        .\WinSCP.exe /log="C:\iSAPScripts\DataTrackerPhase1\WinSCP_get.log" /ini=nul /script=$SFTPGetFilesCMD
        Write-Host "Getting data.csv from SFTP Server..." -ForegroundColor Magenta
        Sleep 2
        Write-Host "Download completed!!" -ForegroundColor Magenta
        sleep 2
        Write-Host "Getting ready in.." -ForegroundColor Magenta
        Write-Host "3" -ForegroundColor red
        sleep 1
        Write-Host "2" -ForegroundColor yellow
        sleep 1
        Write-Host "1" -ForegroundColor green
        sleep 1

    }
    Catch
    {
        Add-Content $log "----------------------------------------------------------------------------------------"
        Add-Content $log "[ERROR]`t Could not connect to SFTP Server. Script will stop!"
        Add-Content $log "[ERROR]`t : $($_.Exception.Message)`r`n"
        Add-Content $log "----------------------------------------------------------------------------------------"
        
    }
}

#------------------------------------------------------
$today = Get-Date; $x =[string]$today; $x = $x.Substring(0,10); $today = $x.replace("`/" ,"-")
$wd = "C:\iSAPScripts\DataTrackerPhase1"
$todayFilePath = "$wd\Report$today.csv"
cd $wd
#----------------------------------------------------
function getUnique{
cd $Wd
Write-Host "[INFO] Getting Unique Records" -ForegroundColor Green
$data = import-csv "$wd\Phase1.csv"
####Get Unique values
$uniqueValues = @()
$array = @()
    forEach ($record in $data) {
    $id =  $record."Personal Reference:People"


        if($id -notin $uniqueValues -and $record."Means of Contact:People" -eq "User e-mail address"){
            Write-Host $id -ForegroundColor Yellow
            $uniqueValues += $id
            $name =$record."Forename1:People"
            $surname = $record."Surname:People"
            $startDate = $record."Start Date:Cont"
            $positionStatus = $record."Position Status"
            $birthDate = $record."Birth Date:People"
            $sex = $record."Sex:People"
            $postcode = $record."Line6:People Address"
            $position = $record."Position"
            $email = $record."Contact At:People"
            $meansofContact = "User e-mail address"
            $unit = $record."Reporting Unit"


            
            
            $row = New-Object Object
            $row | Add-Member -MemberType NoteProperty -Name "EmployeeID" -Value $id
            $row | Add-Member -MemberType NoteProperty -Name "Name" -Value $name
            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $surname
            $row | Add-Member -MemberType NoteProperty -Name "Job Title" -Value $position
            $row | Add-Member -MemberType NoteProperty -Name "Reporting Unit" -Value $unit
            $row | Add-Member -MemberType NoteProperty -Name "StartDate" -Value $startDate
            $row | Add-Member -MemberType NoteProperty -Name "Position Status" -Value $positionStatus
            $row | Add-Member -MemberType NoteProperty -Name "Birth Date" -Value $birthDate
            $row | Add-Member -MemberType NoteProperty -Name "Sex" -Value $sex
            $row | Add-Member -MemberType NoteProperty -Name "Postcode" -Value $postcode
            $row | Add-Member -MemberType NoteProperty -Name "Email" -Value $email
            $array += $row

        }

    }
    $array | Export-Csv -Path "$todayFilePath" -NoTypeInformation 
}

 function 1stPass{
 $lastReportFile = gci $wd/lastReport/ | sort LastWriteTime |Where-Object{$_.name -like "*Report*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.

 $yesterdayFile = Import-Csv  $lastReportFile.FullName
 $todayFile = Import-Csv $todayFilePath
    Write-Host "[INFO]Running 1stPass Function ---> Checking for LEAVERS" -ForegroundColor Green

            $arrayNewRecords = @()
            $arrayToRemove =@()
            forEach ($yesterdayRow in $yesterdayFile) {
                $employeeId = $yesterdayRow.EmployeeID
                $name =$yesterdayRow."Name"
                $surname = $yesterdayRow."Surname"
                $startDate = $yesterdayRow."StartDate"
                $positionStatus = $yesterdayRow."Position Status"
                $birthDate = $yesterdayRow."Birth Date"
                $sex = $yesterdayRow."Sex"
                $postcode = $yesterdayRow."Postcode"
                $position = $yesterdayRow."Job Title"
                $email = $yesterdayRow."Email"
                $unit  = $yesterdayRow."Reporting Unit"


                    #if yesterday USID exist in Absence file today // DO NOTHING
                    if($yesterdayRow.EmployeeID -in $todayFile.EmployeeID){Write-Host "[NO CHANGE] ($employeeId) - $name $surname" -ForegroundColor White -BackgroundColor Blue
                      } #close if


 #if yesterday employeeID NOT in today file it needs to be// LEAVERS
                elseif($yesterdayRow.EmployeeID -notcontains $todayFile.EmployeeID){ 
                    $row2 = New-Object Object
                    ###                      REMOVE RECORD              ######
            $row2 | Add-Member -MemberType NoteProperty -Name "EmployeeID" -Value $employeeId
            $row2 | Add-Member -MemberType NoteProperty -Name "Name" -Value $name
            $row2 | Add-Member -MemberType NoteProperty -Name "Surname" -Value $surname
            $row2 | Add-Member -MemberType NoteProperty -Name "Job Title" -Value $position
            $row2 | Add-Member -MemberType NoteProperty -Name "Reporting Unit" -Value $unit
            $row2 | Add-Member -MemberType NoteProperty -Name "StartDate" -Value $startDate
            $row2 | Add-Member -MemberType NoteProperty -Name "Position Status" -Value $positionStatus
            $row2 | Add-Member -MemberType NoteProperty -Name "Birth Date" -Value $birthDate
            $row2 | Add-Member -MemberType NoteProperty -Name "Sex" -Value $sex
            $row2 | Add-Member -MemberType NoteProperty -Name "Postcode" -Value $postcode
            $row2 | Add-Member -MemberType NoteProperty -Name "Email" -Value $emai
        
            $arrayToRemove += $row2
            
                    Write-Host "[LEAVERS] ($employeeId) - $name $surname" -ForegroundColor Red -BackgroundColor Black 
                } #close elseif
               

                }# ---forEach MAIN
            
                 $LeaversFile = "$wd/Leavers$today.csv"
                 $arrayToRemove | Export-Csv $LeaversFile -NoTypeInformation

            
            }#-- close 1stPass
function 2ndPass{
 $lastReportFile = gci $wd/lastReport/ | sort LastWriteTime |Where-Object{$_.name -like "*Report*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
 $yesterdayFile = Import-Csv  $lastReportFile.FullName
    
    $todayFile = Import-Csv "$todayFilePath" 
    Write-Host "[INFO]Running 2stPass Function ---> Checking for NEW STARTERS" -ForegroundColor Green

    
                $arrayNewRecords = @()
                Foreach ($todayRow in $todayFile) {
                
                $employeeId = $todayRow.EmployeeID

                $name =$todayRow."Name"
                $surname = $todayRow."Surname"
                $startDate = $todayRow."StartDate"
                $positionStatus = $todayRow."Position Status"
                $birthDate = $todayRow."Birth Date"
                $sex = $todayRow."Sex"
                $postcode = $todayRow."Postcode"
                $position = $todayRow."Job Title"
                $email = $todayRow."Email"
                $unit  = $todayRow."Reporting Unit"
                

                                    if($todayRow.EmployeeID -notin $yesterdayFile.EmployeeID) {
                                   $row = New-Object Object
                           ###                      ADD NEW HOLIDAY ABSENCE              ######
                        $row | Add-Member -MemberType NoteProperty -Name "EmployeeID" -Value $employeeId
                        $row | Add-Member -MemberType NoteProperty -Name "Name" -Value $name
                        $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $surname
                        $row | Add-Member -MemberType NoteProperty -Name "Job Title" -Value $position
                        $row | Add-Member -MemberType NoteProperty -Name "Reporting Unit" -Value $unit
                        $row | Add-Member -MemberType NoteProperty -Name "StartDate" -Value $startDate
                        $row | Add-Member -MemberType NoteProperty -Name "Position Status" -Value $positionStatus
                        $row | Add-Member -MemberType NoteProperty -Name "Birth Date" -Value $birthDate
                        $row | Add-Member -MemberType NoteProperty -Name "Sex" -Value $sex
                        $row | Add-Member -MemberType NoteProperty -Name "Postcode" -Value $postcode
                        $row | Add-Member -MemberType NoteProperty -Name "Email" -Value $emai

                        $arrayNewRecords += $row
                        Write-Host "[NEW STARTERS] - ($employeeId) - $name $surname " -ForegroundColor Yellow                 
                    
                            }#--close if

                }
                 $newStarterFile = "$wd/NewStarters$today.csv"
                 $arrayNewRecords | Export-Csv $newStarterFile -NoTypeInformation
Write-Host "----------------------------------------------" -ForegroundColor Green
}#--close 2ndPass

function cleanUP{
Write-Host "[INFO] Starting cleanUp process.." -ForegroundColor Green
cd $wd
Move-Item "Phase1.csv" temp -force
Move-Item "*Report*.csv" "$wd\lastReport\Report$today.csv" -Force
Move-Item "Leavers*.csv" $wd\archive -Force
Move-Item "NewStarters*.csv" $wd\archive -Force
Write-Host "[INFO] cleaned up completed" -ForegroundColor Green
Write-Host "[INFO] Script run successfully" -ForegroundColor Green
}

function sendMail {
Try{
    $x = 0
     $lastReportFile = gci $wd/archive/ | sort LastWriteTime |Where-Object{$_.name -like "*Leave*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
     $Leavers = Import-Csv  $lastReportFile.FullName

     $NewStafftFile = gci $wd/archive/ | sort LastWriteTime |Where-Object{$_.name -like "*Leave*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
     $NewStaff = Import-Csv  $lastReportFile.FullName


    #x $Leavers = Import-CSv $wd\archive\Leavers$today.csv
    #x $NewStaff = Import-Csv $wd\archive\NewStarters$today.csv
    if($Leavers -eq $null -and $NewStaff -eq $null){
            Clear-Host
        Start-Sleep 1
        Write-Host "Sending Email in 3" -ForegroundColor Red
        Start-Sleep 1
        Write-Host "Sending Email in 2" -ForegroundColor Yellow
        Start-Sleep 1
        Write-Host "Sending Email in 1" -ForegroundColor Green
        Write-Host "EMAIL SEND" -ForegroundColor Green
            subEmail-Empty
            Write-Host "[SUCCESS] Email empty report Manager SENT" -ForegroundColor Green
            sleep 2
                While ($x -le 15){ #number of character counts if log file is empty
                $sleep = .8
                Write-Host $art -ForegroundColor Yellow
                sleep $sleep
                Clear-Host
                Write-Host $art -ForegroundColor Green
                sleep $sleep
                Clear-Host
                $x += 1
                
                }
        }

    else{
    subEmail   
    Start-Sleep 1
    Write-Host "Sending Email in 5" -ForegroundColor Red
    Start-Sleep 1
    Write-Host "Sending Email in 4" -ForegroundColor Red
    Start-Sleep 1
    Write-Host "Sending Email in 3" -ForegroundColor Red
    Start-Sleep 1
    Write-Host "Sending Email in 2" -ForegroundColor Yellow
    Start-Sleep 1
    Write-Host "Sending Email in 1" -ForegroundColor Green
    Write-Host ""
    Write-Host ""
    Write-Warning "There have been changes in the structure, with either NEW STARTERS or LEAVERS"
    Write-Host "---------------------------------------------------------------" -ForegroundColor Red 
    Write-host "--------------------------------------------------------------" -foreground Yellow
    Write-Host "Please check the latest report inside the C:/iSAPScipts/DataTrackerPhase1/archive" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "---------------------------------------------------------------" -ForegroundColor Red
    }
}#---close try
Catch{Write-Host "No records found, cant sen't email something went wrong while trying to send the email"}
}

Function subEmail {
Try{
#STATIC VARIABLES for Azure AD
#----------------------------------------------------------
$AZUsername = "itrent.service@fitzroy.org"
$secureString2 = "76492d1116743f0423413b16050a5345MgB8AFcASwBPADcAaQAwAGMAdQBkAEUARgBmAE4AbQA2ADEARgBNAC8ATwBRAFEAPQA9AHwAMQA2AGYAYgBhADgAZQA1AGMANgA1AGUAYgAzADgANAAxADEAYgAzAGMANABlADIANwA5AGUANgBkAGQAZgAyAGEAYgA0ADcAYgBmAGQAMwBmADkANwAzADQANQAyAGQAOQA0ADkAMgAxAGEANQBlAGYAYQBkADEAYgBkADYAZAA2AGMAMwAyADIAYgAyAGEANgBhADQAMQAxAGMAMgA3AGIANgBhADkANwAzAGUAMQA2AGIANwAxAGQAZQBhAGUA"
$password2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR(($secureString2 | ConvertTo-SecureString -Key (Get-Content "C:\iTrentScripts\azure_secret_key.key")))))
$AZPassword = ConvertTo-SecureString -String $password2 -AsPlainText -Force
$AZCredentials = New-Object System.Management.Automation.PSCredential $AZUsername,$AZPassword
#----------------------------------------------------------
     $lastReportFile = gci $wd/archive/ | sort LastWriteTime |Where-Object{$_.name -like "*Leave*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
     $Leavers = Import-Csv  $lastReportFile.FullName

     $NewStafftFile = gci $wd/archive/ | sort LastWriteTime |Where-Object{$_.name -like "*Leave*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
     $NewStaff = Import-Csv  $lastReportFile.FullName

   #x $Leavers = "$wd\archive\Leavers$today.csv" 
   #x $NewStaff = "$wd\archive\NewStarters$today.csv"

    $To = ("paulo.bazzo@fitzroy.org","daniella.ringrose@fitzroy.org","blythe.senior@fitzroy.org")
    $SMTPServer = "smtp.office365.com"
    $SMTPPort = "587"
    $attachment = ($Leavers,$NewStaff)
    $subject = "Changes in Structure Phase 1"
    $body = "Please check files to to identify the employees who are New Starters and Leavers"
    
   Send-MailMessage -From $AZUsername -to $To -Subject $subject `
        -Body $body -SmtpServer $SMTPServer -port $SMTPPort -ErrorAction Stop -Attachments $attachment -UseSsl -Credential $AZCredentials


    }
    Catch
    { Write-Host "[ERROR] Something went wrong while attempting to send the email w/ attachment" -ForegroundColor red}

}



    #sub function located inside Check-Send
Function subEmail-Empty {
Try{
#STATIC VARIABLES for Azure AD
#----------------------------------------------------------
$AZUsername = "itrent.service@fitzroy.org"
$secureString2 = "76492d1116743f0423413b16050a5345MgB8AFcASwBPADcAaQAwAGMAdQBkAEUARgBmAE4AbQA2ADEARgBNAC8ATwBRAFEAPQA9AHwAMQA2AGYAYgBhADgAZQA1AGMANgA1AGUAYgAzADgANAAxADEAYgAzAGMANABlADIANwA5AGUANgBkAGQAZgAyAGEAYgA0ADcAYgBmAGQAMwBmADkANwAzADQANQAyAGQAOQA0ADkAMgAxAGEANQBlAGYAYQBkADEAYgBkADYAZAA2AGMAMwAyADIAYgAyAGEANgBhADQAMQAxAGMAMgA3AGIANgBhADkANwAzAGUAMQA2AGIANwAxAGQAZQBhAGUA"
$password2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR(($secureString2 | ConvertTo-SecureString -Key (Get-Content "C:\iTrentScripts\azure_secret_key.key")))))
$AZPassword = ConvertTo-SecureString -String $password2 -AsPlainText -Force
$AZCredentials = New-Object System.Management.Automation.PSCredential $AZUsername,$AZPassword
#----------------------------------------------------------




    $To = ("paulo.bazzo@fitzroy.org","daniella.ringrose@fitzroy.org","blythe.senior@fitzroy.org")
    $SMTPServer = "smtp.office365.com"
    $SMTPPort = "587"
    $subject = "NO Changes happened yesterday"
    $body = $art
    
   Send-MailMessage -From $AZUsername -to $To -Subject $subject `
        -Body $body -SmtpServer $SMTPServer -port $SMTPPort -ErrorAction Stop -UseSsl -Credential $AZCredentials


    }
    Catch
    { Write-Host "[ERROR] Something went wrong while sending the empty email" -ForegroundColor red}

}


SFTPGetFiles
getUnique
Write-Host "Taking a short break, Im thinking too hard..." -ForegroundColor Magenta
Start-Sleep -Seconds 10
1stPass
2ndPass
cleanUP
sendMail