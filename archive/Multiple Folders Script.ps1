#Multiple Folders Script:


# Base directories
$basePaths = @(
    "(Insert Base File Path 1 Here)",
    "(Insert Base File Path 2 Here)"
)

# Cell numbers to process
$cellNumbers = 1..5

foreach ($base in $basePaths) {

    foreach ($cell in $cellNumbers) {

        # Construct full folder path
        $folderPath = "$base$cell"

        # Check folder exists
        if (Test-Path $folderPath) {
            Write-Host "`nProcessing folder: $folderPath"

            # Process all files in this folder
            Get-ChildItem -Path $folderPath -File | ForEach-Object {

                $file = $_

                # Match channel identifier ch00, ch01, ch02, etc.
                if ($file.Name -match "ch\d{2}") {

                    $channel = $Matches[0]             # Extract channel name e.g. "ch00"
                    $channelFolder = Join-Path $folderPath $channel

                    # Create folder if missing
                    if (-not (Test-Path $channelFolder)) {
                        New-Item -ItemType Directory -Path $channelFolder | Out-Null
                        Write-Host "  Created folder: $channelFolder"
                    }

                    # Move file into channel folder
                    Move-Item -Path $file.FullName -Destination $channelFolder
                    Write-Host "  Moved: $($file.Name) → $channel"
                }
            }
        }
        else {
            Write-Host "`nSkipping missing folder: $folderPath"
        }
    }
}


