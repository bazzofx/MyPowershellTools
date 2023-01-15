Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$outlook = New-Object -com Outlook.Application;
$namespace = $outlook.GetNamespace("MAPI");

Register-ObjectEvent -InputObject $outlook -EventName "AdvancedSearchComplete" -Action {
    Write-Host "ADVANCED SEARCH COMPLETE" $Args.Scope
    Write-Host "Processing, please wait..." -ForegroundColor Yellow

    if ($Args.Results) {  
        foreach ($result in $Args.Results) {
            write-host "=================================================="
            $subject = $result.Subject
            $time = $result.ReceivedTime
            $sender = $result.Sendername
            $body = $result.htmlbody
                $extract = "Subbject: $subject <br><br> Time: $time <br><br> Sender:$sender <br><br> $body" 
            write-host "Subject : $subject"
            write-host "Time: $time"
            write-host "Sender:$sender"
            write-host "=================================================="
            $extract | Out-File "$subject.html"
        }
    
    }


}



Function Get-OutlookInbox($query) {

    $accountsList = $namespace.Folders

    $query = "LOL ITS WORKING"
    $filter = "urn:schemas:httpmail:textdescription LIKE '%"+$query+"%'"

    foreach($account in $accountsList) {
        $scope = $account.FolderPath

        $search = $outlook.AdvancedSearch("'$scope'", $filter, $True)
    }
    

}

Get-OutlookInbox


    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

