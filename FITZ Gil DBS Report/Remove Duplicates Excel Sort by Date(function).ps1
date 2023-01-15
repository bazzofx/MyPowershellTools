$data = @'
Reference,RenewalDate,PropOther1,PropOther2
Ref1,02.03.2022,PVal1_1,PVal2_1
Ref2,03.03.2022,PVal1_2,PVal2_2
Ref3,12.03.2022,PVal1_3,PVal2_3
Ref4,17.03.2022,PVal1_4,PVal2_4
Ref4,18.03.2022,PVal1_4,PVal2_4
Ref1,24.03.2022,PVal1_1,PVal2_1
'@ | ConvertFrom-Csv

$x = $data | Sort-Object Reference, @{ Expression = { Get-Date($_.RenewalDate) }; Descending = $true } | Group-Object -Property Reference

$x | Format-Table # display grouping result

$x | ForEach-Object {
  Write-Host $("Last renewal for " + $_.Name + " was on " + $_.Group[0].RenewalDate)
}