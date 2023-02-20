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
            $object.count +1 # essential to give out last row.
            $Reference = $object.Group[$object.count -1].Reference
            $FirstName = $object.Group[$object.count -1].FirstName
            $Surname = $object.Group[$object.count -1].Surname
            $Unit = $object.Group[$object.count -1].Unit
            $Position = $object.Group[$object.count -1].Position
            $CheckType = $object.Group[$object.count -1].CheckType
            $RenewalDate = $object.Group[$object.count -1].RenewalDate
            $RenewalDate2 = $object.Group[0].RenewalDate
            $backgroundCheck = $object.Group[$object.count -1].BackgroundCheck
            $subs = $object.Group[$object.count -1]."Active DBS Subscription Status"
            $validResult = $object.Group[$object.count -1]."Passed DBS Check?" 
            Write-Host $FirstName $RenewalDate -ForegroundColor Yellow
            Write-Host $RenewalDate2 -ForegroundColor Green
            sleep 1

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
            $object.count +1 # essential to give out last row.
            
            $Reference = $object.Group[$object.count -1].Reference
            $FirstName = $object.Group[$object.count -1].FirstName
            $Surname = $object.Group[$object.count -1].Surname
            $Unit = $object.Group[$object.count -1].Unit
            $Position = $object.Group[$object.count -1].Position
            $CheckType = $object.Group[$object.count -1].CheckType
            $RenewalDate = $object.Group[$object.count -1].RenewalDate
            $backgroundCheck = $object.Group[$object.count -1].BackgroundCheck 


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
sleep 1
Write-Host "Duplicate records removed from Final DBS Check Report" -ForegroundColor Green
Write-Host "Duplicate records removed from Valid DBS Subscription Report" -ForegroundColor Green
