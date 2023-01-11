   $array = @()
   $data = Import-Csv $wd\Combined.csv -Header "Name","OwningProcess","SecurityRisk","State"

    foreach ($x in $data) {
    $name =$x.name
    $id = $x.OwningProcess
    $state = $x.State
            
            if($id-eq 4 -or $id -eq 16800 -or $id -eq 988 -or $id -eq 5404 -or`
            $id -eq 2224 -or $id -eq 84 -or $id -eq 916 -or $id -eq 1008){
            $row = New-Object Object
            Write-Host "Windows Process Safe[id:4]" -ForegroundColor Green
            $row | Add-Member -MemberType NoteProperty -Name "ProcessName" -Value $name
            $row | Add-Member -MemberType NoteProperty -Name "OwningProcess" -Value $id
            $row | Add-Member -MemberType NoteProperty -Name "State"  -Value $state
            $row | Add-Member -MemberType NoteProperty -Name "SecurityRisk"  -Value "Windows Essential"
            }
            elseif ($id -eq 5828){
            $row = New-Object Object
            $row | Add-Member -MemberType NoteProperty -Name "ProcessName" -Value $name
            $row | Add-Member -MemberType NoteProperty -Name "OwningProcess" -Value $id
            $row | Add-Member -MemberType NoteProperty -Name "State"  -Value $state
            $row | Add-Member -MemberType NoteProperty -Name "SecurityRisk"  -Value "Intel Service"
            }
            elseif($id -eq 16216 -or $id -eq 1868 -or $id -eq 1724 ){
            $row = New-Object Object
            $row | Add-Member -MemberType NoteProperty -Name "ProcessName" -Value $name
            $row | Add-Member -MemberType NoteProperty -Name "OwningProcess" -Value $id
            $row | Add-Member -MemberType NoteProperty -Name "State"  -Value $state
            $row | Add-Member -MemberType NoteProperty -Name "SecurityRisk"  -Value "Web Root Services"            
            }

                    
        $array+= $row   
                        }
        $array| Export-Csv $wd\final.csv