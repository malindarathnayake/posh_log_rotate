#-----------------------------------------------
#-------------------- INFOS --------------------
#Script:   poshLogsRotation
#Version:  1.0.1.2
#Date:     27/04/2023
#Author:   Malinda Rathnayake

param (
    [string]$Source
)

if (-not $Source) {
    Write-Host "Please provide a source path using the -Source parameter. Example:"
    Write-Host ".\posh_Rotate.ps1 -Source 'D:\your_log_Path_root'"
    return
}
$MatchedFiles = $null
$ThresholdDate = (Get-Date).AddDays(-30)
$ArchiveFolder = Join-Path -Path $Source -ChildPath "_Archive"
$7ZipExe = "C:\Program Files\7-Zip\7z.exe"  # Adjust the path to your 7zip executable
$filecount = $MatchedFiles.count

# Create the _Archive folder if it doesn't exist
if (-not (Test-Path -Path $ArchiveFolder -PathType Container)) {
    New-Item -Path $ArchiveFolder -ItemType Directory
}

Write-Host "Source Path is valid - $Source"

# Get files older than 30 days
Write-Host "Scanning for files - this may take a few mins"
$MatchedFiles = Get-ChildItem -LiteralPath $Source -Recurse | Where-Object { $_.LastWriteTime -lt $ThresholdDate }

$filecount = $MatchedFiles.count
Write-Host "Completed - $filecount files located"

# Move matched files to the _Archive folder
Write-Host "Creating _Archive if not exist and Moving log folders"
foreach ($file in $MatchedFiles) {
    $RelativePath = $file.FullName.Substring($Source.Length + 1)
    $Destination = Join-Path -Path $ArchiveFolder -ChildPath $RelativePath

    # Create the destination directory if it doesn't exist
    $DestDirectory = [System.IO.Path]::GetDirectoryName($Destination)
    if (-not (Test-Path -Path $DestDirectory -PathType Container)) {
        New-Item -Path $DestDirectory -ItemType Directory -Force
    }
    if (test-path $file.FullName){
    Move-Item -Path $file.FullName -Destination $Destination -Force}
}

Write-Host "Moving log files to $ArchiveFolder Completed"

# Remove empty source folders
Write-Host "Removing empty source folders"

$MovedFolders = $MatchedFiles | ForEach-Object { $_.DirectoryName }
$FoldersToRemove = Get-ChildItem -LiteralPath $Source -Directory | Where-Object { $MovedFolders -contains $_.FullName }

foreach ($folder in $FoldersToRemove) {
    if (Test-Path $folder.FullName) {
        if ((Get-ChildItem -Path $folder.FullName -Recurse | Measure-Object).Count -eq 0) {
            Write-Host "Removing empty source folder $($folder.FullName)"
            Remove-Item -Path $folder.FullName -Recurse -Force
        } else {
            Write-Host "$($folder.FullName) is not empty and will not be removed."
        }
    } else {
        Write-Host "$($folder.FullName) already moved."
    }
}

# Compress each nested folder in _Archive using 7zip with higher compression
$NestedFolders = Get-ChildItem -Path $ArchiveFolder -Directory
foreach ($folder in $NestedFolders) {
    $FolderPath = $folder.FullName
    $ZipFilePath = "$FolderPath.zip"


    # Compress the folder using 7zip with higher compression level (-mx9)
    Write-Host "Processing zip - $ZipFilePath"
    
    Start-Process -FilePath $7ZipExe -ArgumentList "a -tzip -mx9 `"$ZipFilePath`" `"$FolderPath`"" -Wait

    Write-Host "Folder $FolderPath has been compressed to $ZipFilePath to save space"

    # Remove the folder after compression
    Remove-Item -Path $FolderPath -Recurse -Force
    Write-Host "Folder $FolderPath has been removed."
}

Write-Host "Files moved to $ArchiveFolder, nested folders compressed"

# Clean up zip files older than 60 days in _Archive folder.

Write-Host "House cleaning - Remove zip files older than 60 days in _Archive folder"
$logZipFiles = Get-ChildItem -Path $ArchiveFolder -File -Filter *.zip | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-60) }

foreach ($zipFile in $logZipFiles) {
    Remove-Item -Path $zipFile.FullName -Force
    Write-Host "7zip file $($zipFile.FullName) has been removed."
}
