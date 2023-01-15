$counter = 1
$file = "test.txt"
while ($true) {
    $size = $counter * 100KB
    $string = "-".PadRight($size,"-")
    Set-Content -Path $file -Value $string
    Write-Host "File $file is now $size bytes"
    Start-Sleep -Seconds 2
    $counter = $counter * 2
}