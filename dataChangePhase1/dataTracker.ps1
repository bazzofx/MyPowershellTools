#---------------------------------------------------------
$today = Get-Date; $x =[string]$today; $x = $x.Substring(0,10); $today = $x.replace("`/" ,"-")
$wd ="C:\Users\Paulo.Bazzo\OneDrive - FitzRoy\Documents\FitzRoy\PowerShell\MyPowershellTools\dataChangePhase1"
$todayFilePath = "$wd\Report$today.csv"
#--------------------------------------------------------

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

getUnique
Write-Host "Taking a short break, Im thinking too hard..." -ForegroundColor Magenta
Start-Sleep -Seconds 10
1stPass
2ndPass
cleanUP