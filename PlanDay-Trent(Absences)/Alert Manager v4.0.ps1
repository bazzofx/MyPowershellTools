##Connect to SFTP , download report and place inside \.
#This is the main path where the script will live. it must have a folder named "archive" and "logs"

$wd = "C:\Users\Paulo.Bazzo\OneDrive - FitzRoy\Documents\FitzRoy\Trent Projects\In Progress\December\PlanDay - PRJ\To Do\Script Absence Roster\main_v4"

$ErrorActionPreference = "stop"
$DebugPreference = "Continue"
$host.privatedata.VerboseForegroundColor = "Yellow"
$global:sleepTime = 3
$today = get-date -f dd-MM-yyyy
$date = (get-date).AddDays(-1)
$yesterday = $date.ToString("dd-MM-yyyy")
 
$PDFile ="$wd\PDdata.csv"
$RMFile = "$wd\RMdata.csv"
$vlookupFile = "$wd\Combined-$today.csv"
$finalFile = "$wd\Absence Data Merged-$today.csv"

# ------------------ END GLOBAL VARIABLES --------------------------

cd $wd # change directory where files are located

function vLookupManager {
#This will lookup their Full Name and add the Reporting Manager from the RMdata.csv
#Import only SICKNESS AND STATUS that are Approved
     [CmdletBinding()]
     Param([parameter(ValueFromRemainingArguments=$true)][String[]] $args)
#---------- import PlanDate .csv file
$PlanDayData = Import-Csv $PDFile -Header "Date","Title","Salary","FullName","ShiftType","SalaryCode","StartTime","EndTime","Hours","Breaklength","PaidBreakLength","PaidBreak","Department","JobTitle","ShiftStatus","AdministrativeNote","Comment" |
            Where-Object{$_.ShiftStatus -eq "Approved" -or $_StartTime -ne $null -or $EndTime -ne $null} #only imports PlanDay rows that have Shift Type -eq sickness
            
                    Write-Verbose "PlanDay file.csv has been imported successfully"

Try{#---------- import Reporting Manager .csv file
$ReportingManagerData = Import-Csv $RMFile -Header "EmployeeID","FullName","JobTitle","AccessRole","ManagerID","ManagerFullName","PositionManager","OccupancyType","ReportingUnit","EmployeeEmail","ManagerEmail"
                    Write-Verbose "PlanDay file.csv has been imported successfully"
    }#-close try 
Catch{Write-Verbose "Something went wrong trying to import ReportingManager.csv"}

<#set up for Powershell VLookup
The next 3 lines cross reference the FullName on $RMData and $PDData and link the ManagerFullName to it,
it works exactly like a VLOOKUP on excel
#>
    Try{
        $hash=@{} 
        $ReportingManagerData|%{$hash[$_.FullName]=$_."ManagerFullName"}  
        $PlanDayData|select-object Date,FullName,StartTime,EndTime,ShiftType,Hours,@{Name="ManagerFullName"; expression={$hash[$_."Fullname"];$_}}|
        Export-Csv $vlookupFile -NoTypeInformation
        Write-Host "[SUCCESS] VLOOKUP Manager created successfully" -ForegroundColor Green

    }#--close try
    Catch{Write-Host "[ERROR]Something went wrong merging the PDFIle and RMFile together" -ForegroundColor Red}
} # -----------close function //import record if "Approved" or Start/End Time no Null
function vLookupEmployeeID {
#This will lookup their Full Name and add the EmployeeID based on the RMdata.csv
    $PlanDayData = Import-Csv $vlookupFile
$ReportingManagerData = Import-Csv $RMFile -Header "EmployeeID","FullName","JobTitle","AccessRole","ManagerID","ManagerFullName","PositionManager","OccupancyType","ReportingUnit","EmployeeEmail","ManagerEmail"
    
    $hash=@{} 
    $ReportingManagerData|%{$hash[$_.FullName]=$_."EmployeeID"}  
    $PlanDayData|select-object  FullName,`
                                Date,`
                                ShiftType,`
                                StartTime,`
                                EndTime,`
                                Hours,`
                                ManagerFullName,`
                                @{Name="EmployeeID"; expression={$hash[$_."Fullname"];$_}} |    
                                Export-Csv $vlookupFile -NoTypeInformation
    Write-Host "[SUCCESS] VLOOKUP EmployeeID created successfully" -ForegroundColor Green
} # -----------close function
function vLookupEmployeeMail {
#This will lookup their Full Name and add the EmployeeID based on the RMdata.csv
    $PlanDayData = Import-Csv $vlookupFile
$ReportingManagerData = Import-Csv $RMFile -Header "EmployeeID","FullName","JobTitle","AccessRole","ManagerID","ManagerFullName","PositionManager","OccupancyType","ReportingUnit","EmployeeEmail","ManagerEmail"
    
    $hash=@{} 
    $ReportingManagerData|%{$hash[$_.FullName]=$_."EmployeeEmail"}  
    $PlanDayData|select-object  FullName,`
                                Date,`
                                ShiftType,`
                                StartTime,`
                                EndTime,`
                                Hours,`
                                ManagerFullName,`
                                EmployeeID,`
                                @{Name="EmployeeEmail"; expression={$hash[$_."Fullname"];$_}} |    
                                Export-Csv  $vlookupFile -NoTypeInformation
    Write-Host "[SUCCESS] VLOOKUP EmployeeEmail created successfully" -ForegroundColor Green
} # -----------close function
function vLookupManagerEmail {
#This will lookup their Full Name and add the EmployeeID based on the RMdata.csv
    $PlanDayData = Import-Csv $vlookupFile
$ReportingManagerData = Import-Csv $RMFile -Header "EmployeeID","FullName","JobTitle","AccessRole","ManagerID","ManagerFullName","PositionManager","OccupancyType","ReportingUnit","EmployeeEmail","ManagerEmail"
    
    $hash=@{} 
    $ReportingManagerData|%{$hash[$_.FullName]=$_."ManagerEmail"}  
    $PlanDayData|select-object  FullName,`
                                Date,`
                                jobTitle,`
                                ShiftType,`
                                StartTime,`
                                EndTime,`
                                Hours,`
                                ManagerFullName,`
                                EmployeeID,`
                                EmployeeEmail,`
                                @{Name="ManagerEmail"; expression={$hash[$_."Fullname"];$_}} |    
                                Export-Csv  $vlookupFile -NoTypeInformation
    Write-Host "[SUCCESS] VLOOKUP EmployeeEmail created successfully" -ForegroundColor Green
} # -----------close function
function AddUSID {
   #Creates the merged file but now it adds the Unique String Identifier to the row
   #THE USID will help to identify records to ADD/REMOVE/NOT TOUCH 
    [CmdletBinding()]
    Param([parameter(ValueFromRemainingArguments=$true)][String[]] $args)
    #import comnbined Vlookup file
    $data = Import-Csv $vlookupFile 
    $array = @()

    ForEach ($x in $data) {
        $row = New-Object Object
        $employeeID = $x.EmployeeID
        $employeeEmail = $x.EmployeeEmail
        $date = $x.Date
        $fullName = $x.FullName
        $jobTitle = $x.JobTitle
        $startTime = $x.StartTime
        $endTime = $x.EndTime
        $shiftType = $x.ShiftType
        $Hours = $x.Hours
        $manager = $x.ManagerFullName
        $managerEmail = $x.ManagerEmail

        $NamenoSpace = $fullName.Replace(" ","")
        $USID = "USID-" + $date.replace("/","") + $NamenoSpace + "-" + $shiftType + "start"+ $startTime.replace(":","-").substring(0,5) + "end" + $endTime.replace(":","-").substring(0,5)
        Write-Verbose $USID

        $row | Add-Member -MemberType NoteProperty -Name "EmployeeID" -Value $employeeID 
        $row | Add-Member -MemberType NoteProperty -Name "Date" -Value $date
        $row | Add-Member -MemberType NoteProperty -Name "FullName" -Value $fullName
        $row | Add-Member -MemberType NoteProperty -Name "JobTitle" -Value $jobTitle
        $row | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $startTime
        $row | Add-Member -MemberType NoteProperty -Name "EndTime" -Value $endTime
        $row | Add-Member -MemberType NoteProperty -Name "ShiftType" -Value $shiftType
        $row | Add-Member -MemberType NoteProperty -Name "Hours" -Value $Hours
        $row | Add-Member -MemberType NoteProperty -Name "ManagerFullName" -Value $manager
        $row | Add-Member -MemberType NoteProperty -Name "EmployeeEmail" -Value $employeeEmail
        $row | Add-Member -MemberType NoteProperty -Name "ManagerEmail" -Value $managerEmail
        $row | Add-Member -MemberType NoteProperty -Name "USID" -Value $USID

        $array += $row
    } # iteration throught list finished
Try{        
        $array | Export-Csv  $finalFile -NoTypeInformation
        
        Write-Verbose "FetchData Completed"
        Write-Host "[SUCCESS] Added USID to record successfully" -ForegroundColor Green
        Write-Host "[SUCCESS] PlanDay Absence records.csv and Reporting Manager.csv files merged completed" -ForegroundColor Green}
Catch{ Write-Host "[ERROR] Something went wrong merging the 'finalFile' function FetchData" -ForegroundColor Red}

        
} # -----------close function

function MergeReports{
#Currently looked up is done by "Full Name"
vLookupManager
vLookupEmployeeID
vLookupEmployeeMail
vLookupManagerEmail
AddUSID
}



function CreateRAWFiles{

     [CmdletBinding()]
     Param([parameter(ValueFromRemainingArguments=$true)][String[]] $args)
Write-Verbose "Staring CreateRawFiles"
##Export HOLIDAY ABSENCE.csv
Try{
    Import-Csv $finalFile|
    Where-Object {$_.ShiftType -eq "Annual Leave"} |
    Export-Csv $wd/Holiday-RAW-$today.csv -NoTypeInformation
    Write-Host "[SUCCESS] RAW HOLIDAY Report generated" -ForegroundColor Green}
Catch [System.Management.Automation.CmdletInvocationException]{Write-Host "[ERROR] Fail to import .csv file on function CreateRawFiles" -ForegroundColor Black -BackgroundColor Red}

##EXPORT OTHERS ABSENCE.csv
Try {
    Import-Csv $finalFile |
    Where-Object {$_.ShiftType -ne "Sickness" -and $_.ShiftType -ne "Study Leave - Paid" -and $_.ShiftType -ne "Annual Leave" -and $_.StartTime -ne "Start time"} |
    Export-Csv $wd/Other-RAW-$today.csv -NoTypeInformation
    Write-Host "[SUCCESS] RAW SICKNESS absence report generated" -ForegroundColor Green}

Catch [System.Management.Automation.CmdletInvocationException]{Write-Host "[ERROR] Fail to import .csv file on function CreateRawFiles" -ForegroundColor Black -BackgroundColor Red}

##EXPORT SICKNESS ABSENCE.csv
Try{
Import-Csv $finalFile |
    Where-Object {$_.ShiftType -eq "Sickness" -or $_.ShiftType -eq "Study Leave - Paid "} |
    Export-Csv $wd/Sickness-RAW-$today.csv -NoTypeInformation
    Write-Host "[SUCCESS] RAW SICKNESS absence report generated" -ForegroundColor Green}

Catch [System.Management.Automation.CmdletInvocationException]{Write-Host "[ERROR] Fail to import .csv file on function CreateRawFiles" -ForegroundColor Black -BackgroundColor Red}
}

### SUB_FUNCTION ### If Time Hours Absence is Length 5
function timeLength4 {
$timeRow = $global:TimeAbsence

$hours = $timeRow.Substring(0,1)
#Write-host "Hour: $hours"
$minutes = $timeRow.Substring(2,2)
#Write-host "Minutes: $minutes"
$finalTime = "0" + $hours + $minutes
return $finalTime
}

### SUB_FUNCTION ## If Time Hours Absence is Length 5
function timeLength5 {
$timeRow = $global:TimeAbsence

$hours = $timeRow.Substring(0,2)
#Write-Host "Hour: $hours"
$minutes = $timeRow.Substring(3,2)
#Write-host "Minutes: $minutes"
$finalTime = $hours + $minutes
return $finalTime
}


function createLoadFiles {
<# This will create the 4 .csv files that are used for the Batch on Trent
    HolidayAbsences New 
    Other Absences New
    Sickness Absences New
    Remove Absences that are not present anymore
#>
Write-Host "`n--------------- RECORD CHANGES ---------------" -ForegroundColor Green
     Try{
        function FindRecordsHoliday{
        <#For management purposes the 1stPass and 2ndPass are encapsulated inside the FIndRecords function
        1stPass will check if yesterday USID exist in today.csv // DO NOTHING
        1stPass will check if yesterday USID NOT in today USIDyesterday USID needs to be// REMOVE YESTERDAY ABSENCE
        2ndPass will check #if today USID exist but not present in yesterday USID // ADD NEW RECORD ABSENCE
        #>

        $latestReportFile = gci $wd/raw/ | sort LastWriteTime |Where-Object{$_.name -like "*Holiday*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.


        Try{$yesterdayFile = Import-Csv $latestReportFile.FullName}
        Catch {Write-Host "[ERROR] There is not a previous file for Holiday.csv to compare against" -ForegroundColor Red -BackgroundColor Black
        }
        $todayFile = Import-Csv "$wd/*Holiday*.csv"


            function 1stPass{
            $arrayNewRecords = @()
            $global:arrayToRemove =@()
            forEach ($yesterdayRow in $yesterdayFile) {

                    $employeeID = $yesterdayRow.EmployeeID
                    $employeeID = $employeeID.Substring(0,$employeeID.Length-1) #there is a white space after employeeID after export
                    $yesterdayRowUSID = $yesterdayRow.USID
                    $absenceDate = $yesterdayRow.Date
                        $absenceDate+= "xxLOLxx" #this line avoids error if the string is too short
                        $absenceDay = $absenceDate.Substring(0,2)
                        $absenceMonth = $absenceDate.Substring(3,2)
                        $absenceYear = $absenceDate.Substring(6,4)
                        $date = $absenceYear + $absenceMonth + $absenceDay
                    $fullName = $yesterdayRow.FullName
                    $jobTitle = $yesterdayRow.JobTitle
                    $startTime = $yesterdayRow.StartTime
                        $startTimeFormatted = $startTime.Substring(0,$startTime.Length-3).replace(":","")
                    $endTime = $yesterdayRow.EndTime
                        $endTimeFormatted = $endTime.Substring(0,$endTime.Length-3).replace(":","")
                    $shiftType = $yesterdayRow.ShiftType
                    $manager = $yesterdayRow.ManagerFullName
                            $absenceHours = $yesterdayRow.Hours  
                            $global:TimeAbsence = $absenceHours  
                            $timeLength = $absenceHours.Length   #get hours absence lenght of string
                            switch($timeLength){
                            '4' {$formattedTime = timeLength4}   #will check if hour string is 4 or 5 characters
                            '5' {$formattedTime = timeLength5}   #depending on the size of string will format differently
                                            } #--close switch

                #if yesterday USID exist in Absence file today // DO NOTHING
                if($yesterdayRowUSID -in $todayFile.USID){Write-Host "[NO CHANGE] $yesterdayRowUSID" -ForegroundColor White -BackgroundColor Blue
                  #$headers ="PER_REF_NO,ABSENCE_START_DATE,ABSENCE_END_DATE,ABSENCE_START_TYPE,ABSENCE_END_TYPE,ABSENCE_START_TIME,ABSENCE_END_TIME,ABSENCE_TYPE,ABSENCE_START_HOURS,ABSENCE_END_HOURS"|
                  #Out-File "$wd/loadArea/Sickness-TO-LOAD-$today.csv"} #close if ## spit out empty file
                  } #close if
    
                #if yesterday USID NOT in today Absence file it needs to be// REMOVE HOLIDAY ABSENCE
                elseif($yesterdayRowUSID -notcontains $todayFile.USID){ 
                    $row2 = New-Object Object
                    ###                      REMOVE HOLIDAY ABSENCE              ######
                    $row2 | Add-Member -MemberType NoteProperty -Name "PER_REF_NO" -Value $employeeID
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_DATE" -Value $date
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_DATE" -Value $date
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TYPE" -Value "PART"
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TYPE" -Value "PART"
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TIME" -Value $startTimeFormatted
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TIME" -Value $endTimeFormatted
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_TYPE" -Value $shiftType
        
                    $global:arrayToRemove += $row2
            
                    Write-Host "[ABSENCE REMOVED] $yesterdayRowUSID" -ForegroundColor Red -BackgroundColor Black 
                } #close elseif

            }#close forEach
                    ##Export-CSV '$global:arrayToRemove' will happen when all three functions finish run {FindHolidays,FindAbsences,FindOther}"
            } # ---close 1st Pass function
            function 2ndPass {
                #if today USID exist but not present in yesterday Absence file USID // ADD NEW RECORD ABSENCE
        
                $arrayNewRecords = @()
                Foreach ($todayRow in $todayFile) {
                    $employeeID = $todayRow.EmployeeID
                    $employeeID = $employeeID.Substring(0,$employeeID.Length-1) #there is a white space after employeeID after export	 
                    $absenceDate = $todayRow.Date
                        $absenceDate+= "xxLOLxx" #this line avoids error if the string is too short
                        $absenceDay = $absenceDate.Substring(0,2)
                        $absenceMonth = $absenceDate.Substring(3,2)
                        $absenceYear = $absenceDate.Substring(6,4)
                        $date = $absenceYear + $absenceMonth + $absenceDay
                    $startTime = $todayRow.StartTime
                        $startTimeFormatted = $startTime.Substring(0,$startTime.Length-3).replace(":","")
                    $endTime = $todayRow.EndTime
                        $endTimeFormatted = $endTime.Substring(0,$endTime.Length-3).replace(":","")
                    $shiftType = $todayRow.ShiftType
                            $absenceHours = $todayRow.Hours  
                            $global:TimeAbsence = $absenceHours  
                            $timeLength = $absenceHours.Length   #get hours absence lenght of string
                            switch($timeLength){
                            '4' {$formattedTime = timeLength4}   #will check if hour string is 4 or 5 characters
                            '5' {$formattedTime = timeLength5}   #depending on the size of string will format differently
                                            } #--close switch

                    $rowUSID = $todayRow.USID
                    if($todayRow.USID -notin $yesterdayFile.USID) {
                                   $row = New-Object Object
                           ###                      ADD NEW HOLIDAY ABSENCE              ######
                        $row | Add-Member -MemberType NoteProperty -Name "PER_REF_NO" -Value $employeeID
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_DATE" -Value $date
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_DATE" -Value $date
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TYPE" -Value "PART"
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TYPE" -Value "PART"
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TIME" -Value $startTimeFormatted
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TIME" -Value $endTimeFormatted
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_TYPE" -Value "Personal Holiday"
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_HOURS" -Value $formattedTime
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_HOURS" -Value $formattedTime

                        $arrayNewRecords += $row
                        Write-Host "[ADD NEW Annual Leave] - $rowUSID" -ForegroundColor Yellow
                        $reportFile = "$wd/loadArea/Holiday-TO-LOAD-$today.csv"
                        $arrayNewRecords | Export-Csv $reportFile -NoTypeInformation
                            #without the below code the export would have double quotes between records, making the upload fail
                        $data = Get-Content $reportFile
                        $data.Replace('","',",").TrimStart('"').TrimEnd('"') | Out-File $reportFile -Force -Confirm:$false
                    
                    
                            }#--close if
                                }# --close forEach
        } # -----------close 2nd Pass function


        1stPass
        2ndPass
        }

        function FindSicknesssAbsences{
        <#For management purposes the 1stPass and 2ndPass are encapsulated inside the FIndRecords function
        1stPass will check if yesterday USID exist in today.csv // DO NOTHING
        1stPass will check if yesterday USID NOT in today USIDyesterday USID needs to be// REMOVE YESTERDAY ABSENCE
        2ndPass will check #if today USID exist but not present in yesterday USID // ADD NEW RECORD ABSENCE
        #>

        $latestReportFile = gci $wd/raw/ | sort LastWriteTime |Where-Object{$_.name -like "*Sickness*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
        Try{
        $yesterdayFile = Import-Csv $latestReportFile.FullName}
        Catch {Write-Host "[ERROR] There is not a previous file for Sickness.csv to compare against" -ForegroundColor Red -BackgroundColor Black
        }
        $todayFile = Import-Csv "$wd/*Sickness*.csv"


            function 1stPass{
            $arrayNewRecords = @()
            forEach ($yesterdayRow in $yesterdayFile) {

                    $employeeID = $yesterdayRow.EmployeeID
                    $employeeID = $employeeID.Substring(0,$employeeID.Length-1) #there is a white space after employeeID after export
                    $yesterdayRowUSID = $yesterdayRow.USID
                    $absenceDate = $yesterdayRow.Date
                        $absenceDate+= "xxLOLxx" #this line avoids error if the string is too short
                        $absenceDay = $absenceDate.Substring(0,2)
                        $absenceMonth = $absenceDate.Substring(3,2)
                        $absenceYear = $absenceDate.Substring(6,4)
                        $date = $absenceYear + $absenceMonth + $absenceDay    
                    $fullName = $yesterdayRow.FullName
                    $jobTitle = $yesterdayRow.JobTitle
                    $startTime = $yesterdayRow.StartTime
                        $startTimeFormatted = $startTime.Substring(0,$startTime.Length-3).replace(":","")
                    $endTime = $yesterdayRow.EndTime
                        $endTimeFormatted = $endTime.Substring(0,$endTime.Length-3).replace(":","")
                    $shiftType = $yesterdayRow.ShiftType
                    $manager = $yesterdayRow.ManagerFullName
                            $absenceHours = $yesterdayRow.Hours  
                            $global:TimeAbsence = $absenceHours  
                            $timeLength = $absenceHours.Length   #get hours absence lenght of string
                            switch($timeLength){
                            '4' {$formattedTime = timeLength4}   #will check if hour string is 4 or 5 characters
                            '5' {$formattedTime = timeLength5}   #depending on the size of string will format differently
                                            } #--close switch

                #if yesterday USID exist in Absence file today // DO NOTHING
                if($yesterdayRowUSID -in $todayFile.USID){Write-Host "[NO CHANGE] $yesterdayRowUSID" -ForegroundColor White -BackgroundColor Blue
                  #$headers ="PER_REF_NO,ABSENCE_START_DATE,ABSENCE_END_DATE,ABSENCE_START_TYPE,ABSENCE_END_TYPE,ABSENCE_START_TIME,ABSENCE_END_TIME,ABSENCE_TYPE,ABSENCE_START_HOURS,ABSENCE_END_HOURS"|
                  #Out-File "$wd/loadArea/Sickness-TO-LOAD-$today.csv"} #close if ## spit out empty file
                  }
    
                #if yesterday USID NOT in today Absence file it needs to be// REMOVE SICKNESS ABSENCE
                elseif($yesterdayRowUSID -notcontains $todayFile.USID){ 
                    $row2 = New-Object Object
                    ###                      REMOVE SICKNESS ABSENCE              ######
                    $row2 | Add-Member -MemberType NoteProperty -Name "PER_REF_NO" -Value $employeeID
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_DATE" -Value $date
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_DATE" -Value $date
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TYPE" -Value "PART"
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TYPE" -Value "PART"
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TIME" -Value $startTimeFormatted
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TIME" -Value $endTimeFormatted
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_TYPE" -Value $shiftType
        
                    $global:arrayToRemove += $row2  
            
                    Write-Host "[ABSENCE REMOVED] $yesterdayRowUSID" -ForegroundColor Red -BackgroundColor Black 
                } #close elseif

            }#close forEach
                    ##Export-CSV '$global:arrayToRemove' will happen when all three functions finish run {FindHolidays,FindAbsences,FindOther}"
            } # ---close 1st Pass function

            function 2ndPass {
                #if today USID exist but not present in yesterday Absence file USID // ADD NEW SICKNESS ABSENCE
        
                $arrayNewRecords = @()
                Foreach ($todayRow in $todayFile) {
                    $employeeID = $todayRow.EmployeeID	
                    $employeeID = $employeeID.Substring(0,$employeeID.Length-1) #there is a white space after employeeID after export
                    $absenceDate = $todayRow.Date
                        $absenceDate+= "xxLOLxx" #this line avoids error if the string is too short
                        $absenceDay = $absenceDate.Substring(0,2)
                        $absenceMonth = $absenceDate.Substring(3,2)
                        $absenceYear = $absenceDate.Substring(6,4)
                        $date = $absenceYear + $absenceMonth + $absenceDay
                    $startTime = $todayRow.StartTime
                        $startTimeFormatted = $startTime.Substring(0,$startTime.Length-3).replace(":","")
                    $endTime = $todayRow.EndTime
                        $endTimeFormatted = $endTime.Substring(0,$endTime.Length-3).replace(":","")
                    $shiftType = $todayRow.ShiftType
                            $absenceHours = $todayRow.Hours  
                            $global:TimeAbsence = $absenceHours  
                            $timeLength = $absenceHours.Length   #get hours absence lenght of string
                            switch($timeLength){
                            '4' {$formattedTime = timeLength4}   #will check if hour string is 4 or 5 characters
                            '5' {$formattedTime = timeLength5}   #depending on the size of string will format differently
                                            } #--close switch

                    $rowUSID = $todayRow.USID
                    if($todayRow.USID -notin $yesterdayFile.USID) {
                                   $row = New-Object Object
                        ###                      ADD NEW SICKNESS ABSENCE              ######
                        $row | Add-Member -MemberType NoteProperty -Name "PER_REF_NO" -Value $employeeID
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_DATE" -Value $date
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_DATE" -Value $date
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TYPE" -Value "PART"
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TYPE" -Value "PART"
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TIME" -Value $startTimeFormatted
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TIME" -Value $endTimeFormatted
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_TYPE" -Value $shiftType
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_HOURS" -Value $formattedTime
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_HOURS" -Value $formattedTime
        
                        $arrayNewRecords += $row
                        Write-Host "[ADD NEW Sickness Absence] - $rowUSID" -ForegroundColor Magenta
                        $reportFile = "$wd/loadArea/Sickness-TO-LOAD-$today.csv"
                
                        $arrayNewRecords | Export-Csv $reportFile -NoTypeInformation
                            #without the below code the export would have double quotes between records, making the upload fail
                        $data = Get-Content $reportFile
                        $data.Replace('","',",").TrimStart('"').TrimEnd('"') | Out-File $reportFile -Force -Confirm:$false
                    
                    
                            }#--close if
                                }# --close forEach
        } # -----------close 2nd Pass function


        1stPass
        2ndPass
        }

        function FindOtherAbsences{
        $latestReportFile = gci $wd/raw/ | sort LastWriteTime |Where-Object{$_.name -like "*Other*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
        Try{$yesterdayFile = Import-Csv $latestReportFile.FullName}
        Catch {Write-Host "[ERROR] There is not a previous file for Other.csv to compare against" -ForegroundColor Red -BackgroundColor Black
        }
        $todayFile = Import-Csv "$wd/*Other*.csv"

        Try{
            function 1stPass{
            $arrayNewRecords = @()
            forEach ($yesterdayRow in $yesterdayFile) {

                    $employeeID = $yesterdayRow.EmployeeID
                    $employeeID = $employeeID.Substring(0,$employeeID.Length-1) #there is a white space after employeeID after export
                    $yesterdayRowUSID = $yesterdayRow.USID
                    $absenceDate = $yesterdayRow.Date
                        $absenceDate+= "xxLOLxx" #this line avoids error if the string is too short
                        $absenceDay = $absenceDate.Substring(0,2)
                        $absenceMonth = $absenceDate.Substring(3,2)
                        $absenceYear = $absenceDate.Substring(6,4)
                        $date = $absenceYear + $absenceMonth + $absenceDay    
                    $fullName = $yesterdayRow.FullName
                    $jobTitle = $yesterdayRow.JobTitle
                    $startTime = $yesterdayRow.StartTime
                        $startTimeFormatted = $startTime.Substring(0,$startTime.Length-3).replace(":","")
                    $endTime = $yesterdayRow.EndTime
                        $endTimeFormatted = $endTime.Substring(0,$endTime.Length-3).replace(":","")
                    $shiftType = $yesterdayRow.ShiftType
                    $manager = $yesterdayRow.ManagerFullName
                            $absenceHours = $yesterdayRow.Hours  
                            $global:TimeAbsence = $absenceHours  
                            $timeLength = $absenceHours.Length   #get hours absence lenght of string
                            switch($timeLength){
                            '4' {$formattedTime = timeLength4}   #will check if hour string is 4 or 5 characters
                            '5' {$formattedTime = timeLength5}   #depending on the size of string will format differently
                                            } #--close switch
                #if yesterday USID exist in Absence file today // DO NOTHING
                if($yesterdayRowUSID -in $todayFile.USID){Write-Host "[NO CHANGE] $yesterdayRowUSID" -ForegroundColor White -BackgroundColor Blue
                  #$headers ="PER_REF_NO,ABSENCE_START_DATE,ABSENCE_END_DATE,ABSENCE_START_TYPE,ABSENCE_END_TYPE,ABSENCE_START_TIME,ABSENCE_END_TIME,ABSENCE_TYPE,ABSENCE_START_HOURS,ABSENCE_END_HOURS"|
                  #Out-File "$wd/loadArea/Other-TO-LOAD-$today.csv" #spit out empty file
                   
                } #close if
    
                #if yesterday USID NOT in today Absence file it needs to be// REMOVE OTHER ABSENCE
                elseif($yesterdayRowUSID -notcontains $todayFile.USID){ 
                    $row2 = New-Object Object
                        #####                    REMOVE OTHER ABSENCE                   #######
                    $row2 | Add-Member -MemberType NoteProperty -Name "PER_REF_NO" -Value $employeeID
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_DATE" -Value $date
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_DATE" -Value $date
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TYPE" -Value "PART"
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TYPE" -Value "PART"
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TIME" -Value $startTimeFormatted
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TIME" -Value $endTimeFormatted
                    $row2 | Add-Member -MemberType NoteProperty -Name "ABSENCE_TYPE" -Value $shiftType
        
                    $global:arrayToRemove += $row2   
            
                    Write-Host "[ABSENCE REMOVED] $yesterdayRowUSID" -ForegroundColor Red -BackgroundColor Black -ErrorAction Stop 
                } #close elseif

            }#close forEach
                    ##Export-CSV '$global:arrayToRemove' will happen when all three functions finish run {FindHolidays,FindAbsences,FindOther}"
            } # ---close 1st Pass function
            }
        Catch {Write-Host "[ERROR] There is not a previous file for Other.csv to compare against" -ForegroundColor Red -BackgroundColor Black}

        Try{
            function 2ndPass {
                #if today USID exist but not present in yesterday Absence file USID // ADD NEW OTHER ABSENCE
        
                $arrayNewRecords = @()
                Foreach ($todayRow in $todayFile) {
                    $employeeID = $todayRow.EmployeeID	
                    $employeeID = $employeeID.Substring(0,$employeeID.Length-1) #there is a white space after employeeID after export
                    $absenceDate = $todayRow.Date
                        $absenceDate+= "xxLOLxx" #this line avoids error if the string is too short
                        $absenceDay = $absenceDate.Substring(0,2)
                        $absenceMonth = $absenceDate.Substring(3,2)
                        $absenceYear = $absenceDate.Substring(6,4)
                        $date = $absenceYear + $absenceMonth + $absenceDay
                    $startTime = $todayRow.StartTime
                        $startTimeFormatted = $startTime.Substring(0,$startTime.Length-3).replace(":","")
                    $endTime = $todayRow.EndTime
                        $endTimeFormatted = $endTime.Substring(0,$endTime.Length-3).replace(":","")
                    $shiftType = $todayRow.ShiftType
                            $absenceHours = $todayRow.Hours  
                            $global:TimeAbsence = $absenceHours  
                            $timeLength = $absenceHours.Length   #get hours absence lenght of string
                            switch($timeLength){
                            '4' {$formattedTime = timeLength4}   #will check if hour string is 4 or 5 characters
                            '5' {$formattedTime = timeLength5}   #depending on the size of string will format differently
                                            } #--close switch

                    $rowUSID = $todayRow.USID
                    if($todayRow.USID -notin $yesterdayFile.USID) {
                                   $row = New-Object Object
                        #####                         ADD NEW OTHER ABSENCE                   #######
                        $row | Add-Member -MemberType NoteProperty -Name "PER_REF_NO" -Value $employeeID
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_DATE" -Value $date
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_DATE" -Value $date
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TYPE" -Value "PART"
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TYPE" -Value "PART"
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_TIME" -Value $startTimeFormatted
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_TIME" -Value $endTimeFormatted
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_TYPE" -Value $shiftType
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_START_HOURS" -Value $formattedTime
                        $row | Add-Member -MemberType NoteProperty -Name "ABSENCE_END_HOURS" -Value $formattedTime
        
                        $arrayNewRecords += $row
                        Write-Host "[ADD NEW Other Absence] - $rowUSID" -ForegroundColor Cyan
                        $reportFile = "$wd/loadArea/Other-TO-LOAD-$today.csv"
                
                        $arrayNewRecords | Export-Csv $reportFile -NoTypeInformation
                            #without the below code the export would have double quotes between records, making the upload fail
                        $data = Get-Content $reportFile
                        $data.Replace('","',",").TrimStart('"').TrimEnd('"') | Out-File $reportFile -Force -Confirm:$false
                    
                    
                            }#--close if
                                }# --close forEach
        } # -----------close 2nd Pass function
            }
        Catch {Write-Host ""}


        1stPass
        2ndPass
        }

        function ExportAbsencesToRemove{
        $reportFile = "$wd/loadArea/Absences-TO-REMOVE-$today.csv"
        $global:arrayToRemove | Export-Csv $reportFile -NoTypeInformation
        $data = Get-Content $reportFile
        Start-Sleep 5
        $data.Replace('","',",").TrimStart('"').TrimEnd('"') | Out-File $reportFile -Force -Confirm:$false

        }

        $latestReportFile = gci $wd/raw/ | sort LastWriteTime |Where-Object{$_.name -like "*Holiday*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
            if(!($latestReportFile -eq $null)) {FindRecordsHoliday} #if Holiday.csv exist then compare and get new files
            else{Write-Host "[ERROR] File Holiday.csv dont exist inside RAW, comparisson cant be completed." -ForegroundColor Black -BackgroundColor Red}

        $latestReportFile = gci $wd/raw/ | sort LastWriteTime |Where-Object{$_.name -like "*Sickness*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
            if(!($latestReportFile -eq $null)) {FindSicknesssAbsences} #if Sickness.csv exist then compare and get new files
            else{Write-Host "[ERROR] File Sickness.csv dont exist inside RAW, comparisson cant be completed." -ForegroundColor Black -BackgroundColor Red}
        
        $latestReportFile = gci $wd/raw/ | sort LastWriteTime |Where-Object{$_.name -like "*Other*"}| select -last 1 #This is in case the Batch does not run, and its a fail safe mechanism to find the latest file on the folder.
            if(!($latestReportFile -eq $null)) {FindOtherAbsences
                                                ExportAbsencesToRemove
                                                } #if HOther.csv exist then compare and get new files and also get absences to remove.
            else{Write-Host "[ERROR] File Other.csv dont exist inside RAW, comparisson cant be completed." -ForegroundColor Black -BackgroundColor Red}
        } ## close try
     Catch{} ## close catch
     Finally{Write-Host "`n--------------- RECORD CHANGES END ---------------" -ForegroundColor Green
     } # close finally
}#--- close createLoadFiles func



function cleanUp{
Write-Host "Getting ready to clean up files..." -ForegroundColor Yellow
Start-Sleep -Seconds $global:sleepTime
Try{
Remove-Item $vlookupFile #delete previous copy of csv without the USID unecessary .csv
#Move-Item $RMFile $wd/raw/ReportingManagerRaw-$today.csv # MAIN_CORE FILE from SAP Trent - daily download from SFTP
#Move-Item $PDFile $wd/raw/PlanDayRaw-$today.csv # MAIN_CORE FILE from PlanDay - daily download from SFTP
Move-Item $finalFile "$wd/archive"-Force
Move-Item *RAW*.csv "$wd/raw" -Force
Write-Host "[SUCCESS] CleanUp function run completed" -ForegroundColor Green}
Catch{Write-Host "[ERROR] Something went wrong while attempting to move the files into the folder cleanUp function"  -BackgroundColor Red -ForegroundColor black}
}

function run {
MergeReports
CreateRAWFiles
Start-Sleep -Seconds 2
createLoadFiles
Start-Sleep -Seconds 5
cleanup
}
run