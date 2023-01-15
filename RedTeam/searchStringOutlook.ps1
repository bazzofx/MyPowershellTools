$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
$inbox = $namespace.Folders

$searchString = "example string"
$Term = 'pass' 
$emails = $inbox.Items.Restrict($Term)

$emails.Results
