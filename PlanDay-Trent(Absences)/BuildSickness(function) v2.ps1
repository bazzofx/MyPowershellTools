$wd = "C:\Users\Paulo.Bazzo\OneDrive - FitzRoy\Documents\FitzRoy\Trent Projects\In Progress\December\PlanDay - PRJ\To Do\Script Absence Roster\main_v4"
$path = "$wd/PDData.csv"
$data = Import-Csv $path -Header "Date","Title","Salary","Name","ShiftType","SalaryCode","StartTime","EndTime","Hours","Breaklength","PaidBreakLength","PaidBreak","Department","JobTitle","ShiftStatus","AdministrativeNote","Comment"
$array = @()
$count = 0
$dataSize = $data.Count
$global:array = @()
cd $wd 
<#
1stLine runs first, it export the first record Employee A START Sickness
then checks if next row has the same name as previous role
if has same name, skip and check next row until it finds a row with a diferent name.
When it finds a row with a different name, export previous row, Employee A END Sickness
Then it pushes back to 1stLine function where it will export the first record Employee B Start Sickness,
the loops then continues until there are no more rows to check in the file.
#>
function loopCheck{
      [CmdletBinding()]
     Param([parameter(ValueFromRemainingArguments=$true)][String[]] $args)
    $global:count += 1
    Start-Sleep -Milliseconds 100
    Write-Verbose "Inside loop count: $global:count"
    #checking if next row is same name as previous role
    if($global:count -le $dataSize){
        if($data.Name[$global:count] -eq $firstName){
        #if next row has same name dont export
                loopCheck}              
        else {
    #if next row has different name export previous row
    $global:count -= 1
    $outName = $data.Name[$global:count]
    $outDate = $data.Date[$global:count]
    Write-Verbose "Outside loop count $global:count"
    Write-Host "End Sickness recorded added for Name:$outName Date:$outDate" -ForegroundColor Red
    
    $row = New-Object Object
    $row | Add-Member -MemberType NoteProperty -Name "Name" -Value $outName
    $row | Add-Member -MemberType NoteProperty -Name "Date" -Value $outDate
    $global:array += $row
    #then export next row
        $global:count += 1
    1stLine
    #now checks if next row has the same name as previous row
    }
    }
    else{}break
    }

#export row and checks if next row has the same name,
function 1stLine{
     [CmdletBinding()]
     Param([parameter(ValueFromRemainingArguments=$true)][String[]] $args)
$name = $data.Name[$global:count]
$date = $data.Date[$global:count]

    $row = New-Object Object
    $row | Add-Member -MemberType NoteProperty -Name "Name" -Value $name
    $row | Add-Member -MemberType NoteProperty -Name "Date" -Value $date
    $global:array += $row
    $firstName = $name
    $firstDate = $date
    Write-Host "Start Sickness recorded added for Name:$Name Date:$Date" -ForegroundColor Green
    loopCheck
    }

function run {
      [CmdletBinding()]
     Param([parameter(ValueFromRemainingArguments=$true)][String[]] $args)
     $global:count = 0


While ($global:count -le $dataSize){   1stLine  }


Write-Host "-------- Date Check Completed-----" -ForegroundColor Green
$global:array
Write-Host "--------------------------------------" -ForegroundColor Green
}
cls
run
