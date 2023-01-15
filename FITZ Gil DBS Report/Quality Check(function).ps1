$data = Import-Csv "C:\Users\Paulo.Bazzo\Documents\Reports\DBS\quality check dbs report.csv"
#$x = $data | Sort-Object Reference, @{ Expression = { $_.RenewalDate }; Descending = $true } | Group-Object -Property Reference
$x = $data | Group-Object -Property Reference | Sort-Object  "RenewalDate","Reference" -Descending




cls
$first.Group[0].FirstName
$first.Group[0].RenewalDate
$first.Group[1].FirstName
$first.Group[1].RenewalDate
$first.Group[2].FirstName
$first.Group[2].RenewalDate






cls
$first = $x[2]  #What Record group to look for
       <# $count = $first.count
        $count = [int]$count
        #>

$first.Group[2].FirstName
$first.Group[2].RenewalDate 


$first.Group[$count-1].FirstName
$first.Group[$count].RenewalDate


