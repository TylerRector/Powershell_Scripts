$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false  # Set to $true if you want to see Word opening files
$folderPath = ""
$files = Get-ChildItem -Path $folderPath -Filter "*.docx"

foreach ($file in $files) {
	$doc = $wordApp.Documents.Open($file.FullName)
    
	# Print document
	$doc.PrintOut()
    
	# Close document
	$doc.Close($false)
}

# Quit Word application
$wordApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp)


