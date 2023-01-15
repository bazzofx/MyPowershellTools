$RealValidReport = "C:\Users\Paulo.Bazzo\Documents\Reports\DBS\Final DBS Report target date 08-24-2019.csv"
$noduplicatesReport = "$userProfile\Documents\Reports\DBS\Final DBS NO DUPLICATES.csv"

$data = Import-Csv $RealValidReport

$x = $data | Sort-Object Reference, @{ Expression = { $_.RenewalDate }; Ascending = $true } | Group-Object -Property Reference

#$x | Format-Table # display grouping result

$array = @()
ForEach ($object in $x){
            $row = New-Object Object

            $Reference = $object.Group[0].Reference
            $FirstName = $object.Group[0].FirstName
            $Surname = $object.Group[0].Surname
            $Unit = $object.Group[0].Unit
            $Position = $object.Group[0].Position
            $CheckType = $object.Group[0].CheckType
            $RenewalDate = $object.Group[0].RenewalDate
            $backgroundCheck = $object.Group[0].BackgroundCheck
            $subs = $object.Group[0]."Active DBS Subscription Status"
            $validResult = $object.Group[0]."Passed DBS Check?" 


            $row | Add-Member -MemberType NoteProperty -Name "Reference" -Value $Reference
            $row | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $FirstName
            $row | Add-Member -MemberType NoteProperty -Name "Surname" -Value $Surname
            $row | Add-Member -MemberType NoteProperty -Name "Unit" -Value $Unit
            $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $Position
            $row | Add-Member -MemberType NoteProperty -Name "CheckType" -Value $CheckType
            $row | Add-Member -MemberType NoteProperty -Name "RenewalDate" -Value $RenewalDate
            $row | Add-Member -MemberType NoteProperty -Name "BackgroundCheck" -Value $backgroundCheck
            $row | Add-Member -MemberType NoteProperty -Name "Active DBS Subscription Status" -Value $subs
            $row | Add-Member -MemberType NoteProperty -Name "Passed DBS Check?" -Value $validResult

            $array += $row
}
            $array | Export-Csv -Path $noduplicatesReport -NoTypeInformation 
