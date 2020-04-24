
$file = (get-Date).toString("yyyy-MM-dd") + "_DDOLiteTools.zip"

$curDir = Get-Location
$files = @()
$files += Get-Item .\Data 
$files += Get-Item .\Save 
$files += Get-Item .\Settings 
$files += Get-Item .\Utils 
$files += Get-Item *.exe
$files += Get-Item *.md

$files | Select-Object -Unique | Compress-Archive -DestinationPath $file -CompressionLevel Optimal -Force