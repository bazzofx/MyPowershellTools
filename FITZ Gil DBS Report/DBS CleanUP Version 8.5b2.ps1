#$noduplicatesReport = "$userProfile\Documents\Reports\DBS\Final DBS NO DUPLICATES.csv"
#$noduplicatesSubscription = "$userProfile\Documents\Reports\DBS\Final Subscription NO DUPLICATES.csv"

$data = Import-Csv $RealValidReport
$data2 = Import-Csv $subscriptionDBSReport
$x = $data | Sort-Object  @{ Expression = { $x.RenewalDate }; Ascending = $true } | Group-Object -Property Reference 
$y = $data2 | Sort-Object Reference, @{ Expression = { $y.RenewalDate }; Ascending = $true } | Group-Object -Property Reference

$array2 = @()
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
            Write-Host "chillout I am thinking..." -foreground Yellow
            $array += $row
}
            $array | Export-Csv -Path $RealValidReport -NoTypeInformation 




ForEach ($object in $y){
            $row = New-Object Object
            
            $Reference = $object.Group[0].Reference
            $FirstName = $object.Group[0].FirstName
            $Surname = $object.Group[0].Surname
            $Unit = $object.Group[0].Unit
            $Position = $object.Group[0].Position
            $CheckType = $object.Group[0].CheckType
            $RenewalDate = $object.Group[0].RenewalDate
            $backgroundCheck = $object.Group[0].BackgroundCheck 


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
            Write-Host "Thinking a little more..." -ForegroundColor Yellow
            $array2 += $row
}
            $array2 | Export-Csv -Path $subscriptionDBSReport -NoTypeInformation 
Clear-Host
Write-Host "Duplicate records removed from Final DBS Check Report" -ForegroundColor Green
Write-Host "Duplicate records removed from Valid DBS Subscription Report" -ForegroundColor Green
