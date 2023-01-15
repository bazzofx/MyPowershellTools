$ErrorActionPreference= 'Stop'
$userProfile = $env:USERPROFILE

$expiredDate = (get-date).AddMonths(-36) # 3 years ago from today
    $x =[string]$expiredDate
    $x = $x.Substring(0,10)
    $x = $x.replace("`/" ,"-")

$expiredDateSubscription = (get-date).AddMonths(-12) # 1 years ago from today
    $y =[string]$expiredDateSubscription
    $y = $y.Substring(0,10)
    $y = $y.replace("`/" ,"-")

$vaildDate = $expiredDate.AddDays(-1)
# --          DATA SOURCE                  --
$dataPath ="$userProfile\Downloads\MainDBS.csv"
$AllEmployeesPath ="$userProfile\Downloads\AllEmployees.csv"
# --          FOLDER REPORTS                  --
$logFolderReports = "$userProfile\Documents\"
$logFolderDBS = "$userProfile\Documents\Reports\"
# ----                    LOCATION REPORTS FILE NAMES             ----
#---------- VALID REPORTS
$subscriptionDBSReport = "$userProfile\Documents\Reports\DBS\Valid DBS subscription After $y.csv"
$CheckDBSReport = "$userProfile\Documents\Reports\DBS\Valid DBS Records After $x.csv"
$UpdaDBSReport =  "$userProfile\Documents\Reports\DBS\Valid DBS Update Report after $x.csv"
$RealValidReport = "$userProfile\Documents\Reports\DBS\Final Report checking on date $x.csv"

#-----------FAIL REPORTS
$subscriptionFAILDBSReport = "$userProfile\Documents\Reports\DBS\FAIL DBS subscription After $y.csv"
$CheckFAILDBSReport = "$userProfile\Documents\Reports\DBS\FAIL DBS Records After $x.csv"
$UpdateFAILDBSReport = "$userProfile\Documents\Reports\DBS\FAIL DBS Update After $x.csv"
$RealExpiredReport = "$userProfile\Documents\Reports\DBS\DBS Expired Report before $x.csv"
$missingReport = "$userProfile\Documents\Reports\DBS\MISSING DBS Records.csv"
#-------------------------------------------------------
Write-Host "Analyzsing the records..." -ForegroundColor Cyan
Try{
New-Item -Path $logFolderReports -ItemType Directory -Name "Reports"
}
Catch{}
Try{

New-Item -Path $logFolderDBS -ItemType Directory -Name "DBS"}
Catch{}

#############################################################################


function GenerateSubscriptonValid {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Update Service Subscription Paid" -and ($_."RenewalDate" -as [DateTime]) -gt $expiredDateSubscription}
 } #-close Try
 Catch{Write-Warning "MainDBS.csv file not found on $userProfile`\Downloads"}

$array = @()

    ForEach($object in $myCsvFile){
 
            $row = New-Object Object
        
            $Surname = $object.'Surname'
            $FirstName = $object.'FirstName'
            $Reference = $object.'Reference'
            $Unit = $object.'Unit'
            $Position = $object.'Position'
            $CheckType = $object.'CheckType'
            $DateDisclousure = $object.'DateDisclousure'
            $RenewalDate = $object.'RenewalDate'
            $backgroundCheck = $object."BackgroundCheck"

            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $Reference
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $Position
            $row | Add-Member -MemberType NoteProperty -Name "CheckType" -Value $CheckType
            $row | Add-Member -MemberType NoteProperty -Name "RenewalDate" -Value $RenewalDate
            $row | Add-Member -MemberType NoteProperty -Name "BackgroundCheck" -Value $backgroundCheck
            $row | Add-Member -MemberType NoteProperty -Name "DBS Subscription" -Value "Valid"

            $array += $row

            #Write-Host "Checking for Expired Records.." -ForegroundColor Yellow
            }      # iteration throught list finished

$array | Export-Csv -Path $subscriptionDBSReport -NoTypeInformation   
}#-- close function

function GenerateCheckValid {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Check" -and ($_."RenewalDate" -as [DateTime]) -gt $expiredDate }
$DBSupdateList = Import-CSV $subscriptionDBSReport
 
 } #-close Try
 Catch{Write-Warning "MainDBS.csv file not found on $userProfile`\Downloads"}

$array = @()

    ForEach($object in $myCsvFile){
 
            $row = New-Object Object
        
            $Surname = $object.'Surname'
            $FirstName = $object.'FirstName'
            $Reference = $object.'Reference'
            $Unit = $object.'Unit'
            $Position = $object.'Position'
            $CheckType = $object.'CheckType'
            $DateDisclousure = $object.'DateDisclousure'
            $RenewalDate = $object.'RenewalDate'
            $backgroundCheck = $object."BackgroundCheck"

            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $Reference
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $Position
            $row | Add-Member -MemberType NoteProperty -Name "CheckType" -Value $CheckType
            $row | Add-Member -MemberType NoteProperty -Name "RenewalDate" -Value $RenewalDate
            $row | Add-Member -MemberType NoteProperty -Name "BackgroundCheck" -Value $backgroundCheck
            $row | Add-Member -MemberType NoteProperty -Name "DBS Subscription" -Value "Not Present"
            #if ($reference -notin $DBSupdateList."Reference"){
            $array += $row#}

            #Write-Host "Checking for Expired Records.." -ForegroundColor Yellow
            }      # iteration throught list finished

$array | Export-Csv -Path $CheckDBSReport -NoTypeInformation   
}#-- close function

function GenerateUpdateValid {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Update Service Check" -and ($_."RenewalDate" -as [DateTime]) -gt $expiredDate }
$DBSupdateList = Import-CSV $subscriptionDBSReport
 
 } #-close Try
 Catch{Write-Warning "MainDBS.csv file not found on $userProfile`\Downloads"}

$array = @()

    ForEach($object in $myCsvFile){
 
            $row = New-Object Object
        
            $Surname = $object.'Surname'
            $FirstName = $object.'FirstName'
            $Reference = $object.'Reference'
            $Unit = $object.'Unit'
            $Position = $object.'Position'
            $CheckType = $object.'CheckType'
            $DateDisclousure = $object.'DateDisclousure'
            $RenewalDate = $object.'RenewalDate'
            $backgroundCheck = $object."BackgroundCheck"

            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $Reference
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $Position
            $row | Add-Member -MemberType NoteProperty -Name "CheckType" -Value $CheckType
            $row | Add-Member -MemberType NoteProperty -Name "RenewalDate" -Value $RenewalDate
            $row | Add-Member -MemberType NoteProperty -Name "BackgroundCheck" -Value $backgroundCheck
            $row | Add-Member -MemberType NoteProperty -Name "DBS Subscription" -Value "Not Present"
            #if ($reference -notin $DBSupdateList."Reference"){
            $array += $row#}

            #Write-Host "Checking for Expired Records.." -ForegroundColor Yellow
            }      # iteration throught list finished

$array | Export-Csv -Path $UpdaDBSReport -NoTypeInformation   
}#-- close function

function GenerateSubscriptonxpired {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Update Service Subscription Paid" -and ($_."RenewalDate" -as [DateTime]) -lt $expiredDateSubscription}
 } #-close Try
 Catch{Write-Warning "MainDBS.csv file not found on $userProfile`\Downloads"}

$array = @()

    ForEach($object in $myCsvFile){
 
            $row = New-Object Object
        
            $Surname = $object.'Surname'
            $FirstName = $object.'FirstName'
            $Reference = $object.'Reference'
            $Unit = $object.'Unit'
            $Position = $object.'Position'
            $CheckType = $object.'CheckType'
            $DateDisclousure = $object.'DateDisclousure'
            $RenewalDate = $object.'RenewalDate'
            $backgroundCheck = $object."BackgroundCheck"

            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $Reference
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $Position
            $row | Add-Member -MemberType NoteProperty -Name "CheckType" -Value $CheckType
            $row | Add-Member -MemberType NoteProperty -Name "RenewalDate" -Value $RenewalDate
            $row | Add-Member -MemberType NoteProperty -Name "Background Check" -Value $backgroundCheck
            $row | Add-Member -MemberType NoteProperty -Name "DBS Subscription" -Value "Out of Date"

            $array += $row

            #Write-Host "Checking for Expired Records.." -ForegroundColor Yellow
            }      # iteration throught list finished

$array | Export-Csv -Path $subscriptionFAILDBSReport -NoTypeInformation   
}#-- close function

function GenerateCheckExpired {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Check" -and ($_."RenewalDate" -as [DateTime]) -lt $expiredDate}

 
 } #-close Try
 Catch{Write-Warning "MainDBS.csv file not found on $userProfile`\Downloads"}

$array = @()

    ForEach($object in $myCsvFile){
 
            $row = New-Object Object
        
            $Surname = $object.'Surname'
            $FirstName = $object.'FirstName'
            $Reference = $object.'Reference'
            $Unit = $object.'Unit'
            $Position = $object.'Position'
            $CheckType = $object.'CheckType'
            $DateDisclousure = $object.'DateDisclousure'
            $RenewalDate = $object.'RenewalDate'
            $backgroundCheck = $object."BackgroundCheck"

            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $Reference
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $Position
            $row | Add-Member -MemberType NoteProperty -Name "CheckType" -Value $CheckType
            $row | Add-Member -MemberType NoteProperty -Name "RenewalDate" -Value $RenewalDate
            $row | Add-Member -MemberType NoteProperty -Name "BackgroundCheck" -Value $backgroundCheck
            if ($reference -notin $DBSupdateList."Reference"){
            $array += $row}

            #Write-Host "Checking for Expired Records.." -ForegroundColor Yellow
            }      # iteration throught list finished

$array | Export-Csv -Path $CheckFAILDBSReport -NoTypeInformation   
}#-- close function

function GenerateUpdateExpired {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Update Service Check" -and ($_."RenewalDate" -as [DateTime]) -lt $expiredDate}

 
 } #-close Try
 Catch{Write-Warning "MainDBS.csv file not found on $userProfile`\Downloads"}

$array = @()

    ForEach($object in $myCsvFile){
 
            $row = New-Object Object
        
            $Surname = $object.'Surname'
            $FirstName = $object.'FirstName'
            $Reference = $object.'Reference'
            $Unit = $object.'Unit'
            $Position = $object.'Position'
            $CheckType = $object.'CheckType'
            $DateDisclousure = $object.'DateDisclousure'
            $RenewalDate = $object.'RenewalDate'
            $backgroundCheck = $object."BackgroundCheck"

            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $Reference
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $Position
            $row | Add-Member -MemberType NoteProperty -Name "CheckType" -Value $CheckType
            $row | Add-Member -MemberType NoteProperty -Name "RenewalDate" -Value $RenewalDate
            $row | Add-Member -MemberType NoteProperty -Name "BackgroundCheck" -Value $backgroundCheck
            if ($reference -notin $DBSupdateList."Reference"){
            $array += $row}

            #Write-Host "Checking for Expired Records.." -ForegroundColor Yellow
            }      # iteration throught list finished

$array | Export-Csv -Path $UpdateFAILDBSReport -NoTypeInformation   
}#-- close function


function GetRealReport {
#-           Import VALID REPORTS

$validSubscription = Import-Csv $global:subscriptionDBSReport

$validCheck = Import-Csv $global:CheckDBSReport
$validUpdate = Import-Csv $global:UpdaDBSReport
#-----------FAIL REPORTS
$expiredSubscription = Import-Csv $global:subscriptionFAILDBSReport
$expiredCheck = Import-Csv $global:CheckFAILDBSReport
$expiredUpdate = Import-Csv $global:UpdateFAILDBSReport
#-------------------------------------------------------



Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck"

 
 } #-close Try
Catch{Write-Host "[ERROR] MainDBS.csv file not found on $userprofile`\Downdloads" -ForegroundColor Red}

        $arrayValid = @()
        $arrayExpired = @()
        
    ForEach($object in $myCsvFile){
 
            $row = New-Object Object
        
            $Surname = $object.'Surname'
            $FirstName = $object.'FirstName'
            $Reference = $object.'Reference'
            $Unit = $object.'Unit'
            $Position = $object.'Position'
            $CheckType = $object.'CheckType'
            $DateDisclousure = $object.'DateDisclousure'
            $RenewalDate = $object.'RenewalDate'
            $backgroundCheck = $object."BackgroundCheck"

            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $Reference
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $Position
            $row | Add-Member -MemberType NoteProperty -Name "CheckType" -Value $CheckType
            $row | Add-Member -MemberType NoteProperty -Name "RenewalDate" -Value $RenewalDate
            $row | Add-Member -MemberType NoteProperty -Name "BackgroundCheck" -Value $backgroundCheck
            
            
            if($Reference -in $validUpdate."Reference" -or $Reference -in $validCheck."Reference" -and $Reference -in $validSubscription."Reference"){
            $row | Add-Member -MemberType NoteProperty -Name "Subscription Status" -Value "Valid"
            $row | Add-Member -MemberType NoteProperty -Name "Passed DBS Check?" -Value "YES"
            $arrayValid += $row}
            Elseif($Reference -in $validUpdate."Reference" -or $Reference -in $validCheck."Reference" -and $Reference -notin $validSubscription."Reference"){
            $row | Add-Member -MemberType NoteProperty -Name "Subscription Status" -Value "Not Valid"
            $row | Add-Member -MemberType NoteProperty -Name "Passed DBS Check?" -Value "YES"
            $arrayValid += $row}
            Elseif($Reference -in $expiredCheck."Reference" -and $Reference -notin $validCheck."Reference" -and $Reference -notin $validUpdate."Reference"){
            $row | Add-Member -MemberType NoteProperty -Name "Subscription Status" -Value "Not Valid"        
            $row | Add-Member -MemberType NoteProperty -Name "Passed DBS Check?" -Value "NO"
            $arrayValid += $row}
            Elseif ($Reference -in $expiredUpdate."Reference" -and $Reference -notin $validCheck."Reference" -and $Reference -notin $validUpdate."Reference"){
            $row | Add-Member -MemberType NoteProperty -Name "Subscription Status" -Value "Not Valid"          
            $row | Add-Member -MemberType NoteProperty -Name "Passed DBS Check?" -Value "NO"            
            $arrayValid += $row}

} #--- close For Each - Export Real Valid Report

            $arrayValid | Export-Csv -Path $RealValidReport -NoTypeInformation  
            #$arrayExpired | Export-Csv -Path $RealExpiredReport -NoTypeInformation          
            }# -- close function


function allUsersMissingDBS {
Try{
$dataPath ="$userProfile\Downloads\MainDBS.csv"
$allEmployeesReport = Import-Csv $AllEmployeesPath -Header "Reference", "FirstName", "Surname", "Reporting Unit" | ?{$_.Reference -notin $validDBS.Reference -and $_.Reference -notin $expiredDBS.Reference }
$finalReport = Import-Csv $global:RealValidReport
}
Catch{Write-Warning $Error[0]
    Write-Warning "Check if MainDBS.csv and AllEmployees.csv are present found on $userProfile`\Downloads"}

$allEmployeesReport = Import-Csv $AllEmployeesPath -Header "Reference", "FirstName", "Surname", "Reporting Unit" | ?{$_.Reference -notin $finalReport.Reference}



$array5 = @()
        ForEach($record in $allEmployeesReport){

            $ref =  $record.Reference
            $FirstName =  $record.FirstName
            $Surname = $record.Surname
            $Unit = $record."Reporting Unit"

            $row = New-Object Object

            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $ref
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Check Type" -Value "DBS or Enhance DBS"
            $row | Add-Member -MemberType NoteProperty -Name "DBS Number" -Value "MISSING"
            $row | Add-Member -MemberType NoteProperty -Name "Status" -Value "NOT VALID"
            $array5 += $row

}

$array5 | Export-Csv -Path $missingReport -NoTypeInformation 

 
} #--- close function


function HouseKeeping{
$userProfile = $env:USERPROFILE
$logFolderArchive = "$userProfile\Documents\Reports\DBS"
$ArchiveName = "Source $(get-date -f dd-MM-yyy)"
$ArchivePath = $logFolderArchive + "\$ArchiveName"

Try{

New-Item -Path $logFolderArchive -ItemType Directory -Name $ArchiveName}
Catch{}

#mv $subscriptionDBSReport  $ArchivePath
mv $CheckFAILDBSReport $ArchivePath -force
mv $UpdateFAILDBSReport $ArchivePath -force
mv $CheckDBSReport $ArchivePath -force
mv $UpdaDBSReport $ArchivePath -force
mv $subscriptionDBSReport $ArchivePath -force
mv $subscriptionFAILDBSReport $ArchivePath -force



}
clear
sleep 1

GenerateSubscriptonValid
GenerateCheckValid
GenerateUpdateValid

GenerateCheckExpired
GenerateUpdateExpired 
GenerateSubscriptonxpired


GetRealReport
allUsersMissingDBS
HouseKeeping
Write-Host "Your final reports are saved on $logFolderDBS" -ForegroundColor Yellow -BackgroundColor Black    
