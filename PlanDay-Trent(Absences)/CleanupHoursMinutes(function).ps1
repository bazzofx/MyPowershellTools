$data = "4.35", "7.50", "6.00", "0.10","12.50","3.25","9.00"
$data= "8.35", "8.50", "8.11", "0.45","07.25","13.15","0.30"
#$path = "C:\Users\Paulo.Bazzo\OneDrive - FitzRoy\Documents\FitzRoy\Trent Projects\In Progress\December\PlanDay - PRJ\To Do\Script Absence Roster\main_v3\Holiday-RAW-08-01-2023.csv"
$path = "C:\Users\Paulo.Bazzo\OneDrive - FitzRoy\Documents\FitzRoy\Trent Projects\2023\January\PlanDay - PRJ\To Do\Script Absence Roster\PDdata.csv"
cls
<#
This will look at the numbers similar to the $data and check if they are 4 digits or 5 digits,
depending which they are a different function will run timeLenght4 or timeLength5 to transform their data
into something more digestible for Trent
#>

function timeLength4 {
$timeRow = $global:TimeAbsence

$hours = $timeRow.Substring(0,1)
#Write-host "Hour: $hours"
$minutes = $timeRow.Substring(2,2)
#Write-host "Minutes: $minutes"
$finalTime = "0" + $hours + $minutes
return $finalTime
}

## If Time Hours Absence is Length 5
function timeLength5 {
$timeRow = $global:TimeAbsence

$hours = $timeRow.Substring(0,2)
#Write-Host "Hour: $hours"
$minutes = $timeRow.Substring(3,2)
#Write-host "Minutes: $minutes"
$finalTime = $hours + $minutes
return $finalTime
}


$data = Import-Csv $path

foreach($rowFile in $data) {
        $absenceHours = $rowFile.Hours
        $global:TimeAbsence = $absenceHours
        $timeLength = $absenceHours.Length
        switch($timeLength){
        '4' {$formatedTime = timeLength4}
        '5' {$formatedTime = timeLength5}
                        } #--close switch
        $formatedTime 

        
}