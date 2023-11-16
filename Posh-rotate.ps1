function Compress-And-Archive {
    param (
        [string]$Source,
        [int]$DaysThreshold = 30
    )

    try {
        # Validate Source path
        if (-not (Test-Path -Path $Source -PathType Container)) {
            throw "Source path '$Source' does not exist or is not a directory."
        }

        $ThresholdDate = (Get-Date).AddDays(-$DaysThreshold)
        $ArchiveFolder = Join-Path -Path $Source -ChildPath "_Archive"
        $7ZipExe = "C:\Program Files\7-Zip\7z.exe"  # Adjust the path to your 7zip executable

        # Create the _Archive folder if it doesn't exist
        if (-not (Test-Path -Path $ArchiveFolder -PathType Container)) {
            New-Item -Path $ArchiveFolder -ItemType Directory
        }

        # Get files older than the specified days
        $MatchedFiles = Get-ChildItem -LiteralPath $Source -Recurse | Where-Object { $_.LastWriteTime -lt $ThresholdDate }

        if ($MatchedFiles -eq $null) {
            Write-Host "No logs older than $DaysThreshold days found in $Source. Exiting."
            return
        }

        # Move matched files to the _Archive folder
        foreach ($file in $MatchedFiles) {
            $RelativePath = $file.FullName.Substring($Source.Length + 1)
            $Destination = Join-Path -Path $ArchiveFolder -ChildPath $RelativePath

            # Create the destination directory if it doesn't exist
            $DestDirectory = [System.IO.Path]::GetDirectoryName($Destination)
            if (-not (Test-Path -Path $DestDirectory -PathType Container)) {
                New-Item -Path $DestDirectory -ItemType Directory -Force
            }

            Move-Item -Path $file.FullName -Destination $Destination -Force
        }

        Write-Host "Files moved to $ArchiveFolder"

        $MovedFolders = $MatchedFiles | ForEach-Object { $_.DirectoryName }
        $FoldersToRemove = Get-ChildItem -LiteralPath $Source -Directory | Where-Object { $MovedFolders -contains $_.FullName }

        foreach ($folder in $FoldersToRemove) {
            Remove-Item -Path $folder.FullName -Recurse -Force
        }

        # Compress each nested folder in _Archive using 7zip with higher compression
        $NestedFolders = Get-ChildItem -Path $ArchiveFolder -Directory
        foreach ($folder in $NestedFolders) {
            $FolderPath = $folder.FullName
            $ZipFilePath = "$FolderPath.zip"

            # Compress the folder using 7zip with higher compression level (-mx9)
            Start-Process -FilePath $7ZipExe -ArgumentList "a -tzip -mx9 `"$ZipFilePath`" `"$FolderPath`"" -Wait

            Write-Host "Folder $FolderPath has been compressed to $ZipFilePath with higher compression"

            # Remove the folder after compression
            Remove-Item -Path $FolderPath -Recurse -Force
            Write-Host "Folder $FolderPath has been removed."
        }

        Write-Host "Files moved to $ArchiveFolder, nested folders compressed"
    }
    catch {
        Write-Host "Error: $_"
    }
}

# Define Paths to process
$PathsToProcess = @(
    'D:\deviceLogs\',
    'D:\deviceLogs - Copy'
    # Add more paths as needed
)

foreach ($Path in $PathsToProcess) {
    Compress-And-Archive -Source $Path -DaysThreshold 30
}
