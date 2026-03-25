$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.Visible = $true
$folderPath = ""
$files = Get-ChildItem -Path $folderPath -Filter "*.pptx"

foreach ($file in $files) {
	$ppt = $pptApp.Presentations.Open($file.FullName, $false, $false, $false)
	$ppt.PrintOut()
	$ppt.Close()
}

$pptApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp)

