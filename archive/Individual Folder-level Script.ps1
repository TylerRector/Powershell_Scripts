#Individual Folder-level Script:

# Set the folder you want to process
$folderPath = "C:\path\to\your\folder"

# Get all files in the folder
Get-ChildItem -Path $folderPath -File | ForEach-Object {

    $file = $_
    
    # Look for channel pattern "ch00", "ch01", etc.
    if ($file.Name -match "ch\d{2}") {

        # Extract channel identifier
        $channel = $Matches[0]

        # Make directory if it doesn't already exist
        $channelFolder = Join-Path $folderPath $channel
        if (-not (Test-Path $channelFolder)) {
            New-Item -ItemType Directory -Path $channelFolder | Out-Null
        }

        # Move file into the channel folder
        Move-Item -Path $file.FullName -Destination $channelFolder
    }
}
