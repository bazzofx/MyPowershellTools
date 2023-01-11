## Get all process that are talking to the internet
$wd = "C:\Users\Paulo.Bazzo\OneDrive - FitzRoy\Documents\FitzRoy\PowerShell\MyPowershellTools\CheckSuspiciousProcess"

$InternetProcess = Get-NetTCPConnection |
Where-object {$_.State -eq "Listen" -or $_.State -eq "Established Internet"}| Select-Object LocalAddress,`
                              LocalPort,`
                              RemoteAddress,`
                              RemotePort,`
                              State,`
                              AppliedSetting,`
                              OwningProcess |
                             Export-Csv $wd\temp1.csv    #Out-File $wd\temp.txt

Start-Sleep 2
###$data = Get-Content $wd\temp.csv
$data = Import-Csv $wd\temp1.csv 
##Get all process names and process number that were included on the internetProcess and export it
    $array = @()
    $activeProcess = foreach($x in $data){
        Get-Process | Where-Object {$_.id -eq $x.OwningProcess}
            } 
    $activeProcess | Select ProcessName, id |
    Export-Csv  $wd\temp2.csv          
    Start-Sleep 2
#Vlookup
$hash=@{}
$netFile = import-csv "$wd\temp1.csv" -header "LocalAddress","LocalPort","RemoteAddress","RemotePort","State","AppliedSetting","OwningProcess"
$processFile = import-csv "$wd\temp2.csv" -Header "ProcessName","OwningProcess"
$netFile|%{$hash[$_.OwningProcess]=$_."State"}
##combine both files
$processFile|select-object ProcessName,OwningProcess,SecurityRisk,@{name="State"; expression={$hash[$_."OwningProcess"];$_}} |export-csv $wd\Combined.csv
##add row to the file if file is safe
Start-Sleep 2
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
            else{
            $row = New-Object Object
            $row | Add-Member -MemberType NoteProperty -Name "ProcessName" -Value $name
            $row | Add-Member -MemberType NoteProperty -Name "OwningProcess" -Value $id
            $row | Add-Member -MemberType NoteProperty -Name "State"  -Value $state
            $row | Add-Member -MemberType NoteProperty -Name "SecurityRisk"  -Value "Unknown" 
            }

                    
        $array+= $row   
                        }
        $array| Export-Csv $wd\final.csv

##add row to the file short description

##add row to the file if the process is suspicious


