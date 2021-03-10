
$file = (get-Date).toString("yyyy-MM-dd") + "_DDOLiteTools.zip"


#Remove files that cause issues
if (Test-Path .\Save\Main.compendium) { Remove-Item .\Save\Main.compendium -Force -Confirm:$false }
if (Test-Path .\Save\Notes.txt)       { Remove-Item .\Save\Notes.txt       -Force -Confirm:$false }
if (Test-Path .\Save\TEST*.build)     { Remove-Item .\Save\TEST*.build     -Force -Confirm:$false }
if (Test-Path .\Data\Builder\*.bak)   { Remove-Item .\Data\Builder\*.bak   -Force -Confirm:$false }

$curDir = Get-Location
$files = @()
$files += Get-Item .\Data 
$files += Get-Item .\Save 
$files += Get-Item .\Settings 
$files += Get-Item .\Utils 
$files += Get-Item *.exe
$files += Get-Item *.md




$files | Select-Object -Unique | Compress-Archive -DestinationPath $file -CompressionLevel Optimal -Force

Write-Host "Created $file"

