Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
    $outlook = New-Object -com Outlook.Application;
    $namespace = $outlook.GetNamespace("MAPI");


$Term = 'Password'
$Scope = 'Inbox'

$Emails = $outlook.AdvancedSearch( $Scope, "urn:schemas:httpmail:subject LIKE '%$Term%'", $true )

Start-Sleep -Seconds 5 # no sleep will give null results

$Emails.Results | Select-Object -Property Subject