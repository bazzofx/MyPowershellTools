$art = @"
                                        __
                 ,-_                  (`  ).
                 |-_'-,              (     ).
                 |-_'-'           _(        '`.
        _        |-_'/        .=(`(      .     )
       /;-,_     |-_'        (     (.__.:-`-_.'
      /-.-;,-,___|'          `(       ) )
     /;-;-;-;_;_/|\_ _ _ _ _   ` __.:'   )
        x_( __`|_P_|-;-;-;,|        `--'
        |\ \    _||   `-;-;-'
        | \`   -_|.      '-'
        | /   /-_| `
        |/   ,'-_|  \
        /____|'-_|___\
 _..,____]__|_\-_'|_[___,.._
'                          ``'--,..,.      mic
"@
$userProfile = $env:USERPROFILE
$path = "$userProfile\Downloads"
 
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
    $outlook = New-Object -com Outlook.Application;
    $namespace = $outlook.GetNamespace("MAPI");
    $accounts = $namespace.Accounts | Select SmtpAddress | Set-Clipboard # get all accounts on profile save them into clipboard
    $accountsArray = Get-Clipboard


foreach ($acc in $accountsArray){
   $accString = $acc.Substring($acc.IndexOf("=") + 1).Trim("}").ToLower() #cleanup their names
    Write-host $accString -ForegroundColor Yellow

$Term = 'Password' #change the search criteria you are looking for here
$Scope = "'\$accString'"

#the urn:schemas:httpmail:subject searchs for matching Term on body of message
$Emails = $outlook.AdvancedSearch($Scope,"urn:schemas:httpmail:subject LIKE '%$Term%'", $true )


Start-Sleep -Seconds 10 # if it does not sleep well, the results will be BLANK
$Emails.Results | Select-Object -Property Subject,ReceivedTime,Sendername,htmlbody

$file = $path + "\" + $accString + ".html"
$Emails.Results| Out-File $file
Write-Host "Search completed"

}#--close function

#Garbage Collection
    Write-Host "GC collected" -ForegroundColor Yellow
    Remove-Variable outlook
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Try{[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)}
    Catch{}

    Write-Host $art