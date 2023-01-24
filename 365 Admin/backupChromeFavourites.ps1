$today =Get-Date -Format dd/MM/yyyy
$today = $today.Replace("/","-")
$userName = $env:USERNAME
$favFile = "C:\Users\$userName\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"
$destination = "C:\Users\Paulo.Bazzo\OneDrive - FitzRoy\Documents\FitzRoy\PowerShell\MyPowershellTools\365 Admin\favourites\Bookmarks-$today"
copy-item $favFile $destination