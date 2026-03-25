#Root experiment folders
$root = "X:\Path\To\Experiment"

#Folder name pattern
$folderNamePattern = 'Cell\s\d+.*Timelapse'

#Channel identifier pattern
$channelPattern = 'ch\d{2}'

Get-ChildItem -Path $root -Directory | Where-Object {
    $_.Name -match $folderNamePattern
} | ForEach-Object {

    $folderPath = $_.FullName
    Write-Host "`nProcessing folder: $folderPath"

    Get-ChildItem -Path $folderPath -File | ForEach-Object {

        if ($_.Name -match $channelPattern) {

            $channel       = $Matches[0]
            $channelFolder = Join-Path $folderPath $channel

            if (-not (Test-Path $channelFolder)) {
                New-Item -ItemType Directory -Path $channelFolder | Out-Null
                Write-Host "  Created folder: $channelFolder"
            }
            Move-Item -Path $_.FullName -Destination $channelFolder
            Write-Host "  Moved: $($_.Name) → $channel"
        }
    }
}
