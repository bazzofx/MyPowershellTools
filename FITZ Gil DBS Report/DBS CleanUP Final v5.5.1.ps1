$ErrorActionPreference= 'Stop'
$userProfile = $env:USERPROFILE
$expiredDate = (get-date).AddMonths(-12) # 1 years ago from today
$vaildDate = $expiredDate.AddDays(-1)
# --          DATA SOURCE                  --
$dataPath ="$userProfile\Downloads\MainDBS.csv"
$AllEmployeesPath ="$userProfile\Downloads\AllEmployees.csv"
# --          FOLDER REPORTS                  --
$logFolderReports = "$userProfile\Documents\"
$logFolderDBS = "$userProfile\Documents\Reports\"
# ----                    LOCATION REPORTS FILE NAMES             ----
#---------- VALID REPORTS
$subscriptionDBSReport = "$userProfile\Documents\Reports\DBS\Valid DBS subscription After $(get-date -f dd-MM-yyy).csv"
$CheckDBSReport = "$userProfile\Documents\Reports\DBS\Valid DBS Records After $(get-date -f dd-MM-yyy).csv"
$RealValidReport = "$userProfile\Documents\Reports\DBS\Real Valid Report $(get-date -f dd-MM-yyy).csv"
#-----------FAIL REPORTS
$subscriptionFAILDBSReport = "$userProfile\Documents\Reports\DBS\FAIL DBS subscription After $(get-date -f dd-MM-yyy).csv"
$CheckFAILDBSReport = "$userProfile\Documents\Reports\DBS\FAIL DBS Records After $(get-date -f dd-MM-yyy).csv"
$RealExpiredReport = "$userProfile\Documents\Reports\DBS\Real Expired Report $(get-date -f dd-MM-yyy).csv"
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


function GenerateUpdateValid {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Update Service Subscription Paid" -and ($_."RenewalDate" -as [DateTime]) -gt $expiredDate}
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

            $array += $row

            #Write-Host "Checking for Expired Records.." -ForegroundColor Yellow
            }      # iteration throught list finished

$array | Export-Csv -Path $subscriptionDBSReport -NoTypeInformation   
}#-- close function

function GenerateCheckValid {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Check" -and ($_."RenewalDate" -as [DateTime]) -gt $expiredDate}
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
            if ($reference -notin $DBSupdateList."Reference"){
            $array += $row}

            #Write-Host "Checking for Expired Records.." -ForegroundColor Yellow
            }      # iteration throught list finished

$array | Export-Csv -Path $CheckDBSReport -NoTypeInformation   
}#-- close function

function GenerateUpdateExpired {
#Import records who had DBS expire date before $targetDate
Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck" | 
 Where-Object { $_."BackgroundCheck" -eq "DBS Update Service Subscription Paid" -and ($_."RenewalDate" -as [DateTime]) -lt $expiredDate}
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


function GetRealReport {
#-           Import VALID REPORTS
$validUpdate = Import-Csv $global:subscriptionDBSReport
$validCheck = Import-Csv $global:CheckDBSReport
#-----------FAIL REPORTS
$expiredUpdate = Import-Csv $global:subscriptionFAILDBSReport
$expiredCheck = Import-Csv $global:CheckFAILDBSReport
#-------------------------------------------------------



Try{
$myCsvFile =Import-CSV "$dataPath" -Header "Surname", "FirstName", "Reference", "Unit", "JoiningDate", "Position", "CheckType", "DateDisclousure", "RenewalDate", "Latest", "BackgroundCheck"

 
 } #-close Try
Catch{Write-Warning "MainDBS.csv file not found on $userProfile`\Downloads"}

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
            
            
            if($Reference -in $validUpdate."Reference" -and $BackgroundCheck -eq "DBS Update Service Subscription Paid"){
            $arrayValid += $row} # get valid
            Elseif($Reference -in $validCheck."Reference" -and $Reference -notin $validUpdate."Reference" -and $Reference -notin $expiredUpdate."Reference") {
            $arrayValid += $row} # get valid
            Elseif($Reference -in $expiredCheck."Reference" -and $Reference -notin $validUpdate."Reference" -and $Reference -notin $expiredUpdate."Reference"){
            $arrayExpired += $row} # get Expired
            Elseif($Reference -in $expiredCheck."Reference" -and $Reference -in $expiredUpdate."Reference"){
            $arrayExpired += $row} # get Expired


} #--- close For Each - Export Real Valid Report

            $arrayValid | Export-Csv -Path $RealValidReport -NoTypeInformation  
            $arrayExpired | Export-Csv -Path $RealExpiredReport -NoTypeInformation          
            }# -- close function

function HouseKeeping{
$userProfile = $env:USERPROFILE
$logFolderArchive = "$userProfile\Documents\Reports\DBS"
$ArchiveName = "Source $(get-date -f dd-MM-yyy)"
$ArchivePath = $logFolderArchive + "\$ArchiveName"

Try{

New-Item -Path $logFolderArchive -ItemType Directory -Name $ArchiveName}
Catch{Write-Warning $Error[0]}

mv $subscriptionDBSReport  $ArchivePath -Force
mv $CheckDBSReport $ArchivePath -Force
mv $subscriptionFAILDBSReport $ArchivePath -Force
mv $CheckFAILDBSReport $ArchivePath -Force
} # -- move extra reports to Source Folder



function run {
clear
sleep 1

GenerateUpdateValid
GenerateCheckValid
GenerateUpdateExpired 
GenerateCheckExpired


GetRealReport
HouseKeeping

Write-Host "Your final reports are saved on $logFolderArchive" -ForegroundColor Yellow -BackgroundColor Black    
}

run