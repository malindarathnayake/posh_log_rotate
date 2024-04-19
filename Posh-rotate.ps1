#-----------------------------------------------
#-------------------- INFOS --------------------
#Script:   poshLogsRotation
#Version:  1.0.2
#Date:     04/19/2024
#Author:   Malinda Rathnayake

#Variables-----------------------------
param (
    [string]$Source,
    [int]$Thresholddays = 30,
    [int]$ArchiveDays = 60,
    [string]$7ZipExe = "C:\Program Files\7-Zip\7z.exe"
)

if (-not $Source) {
    Write-Host "Please provide a source path using the -Source parameter. Example:"
    Write-Host ".\posh_Rotate.ps1 -Source 'D:\your_log_Path_root'"
    return
}

$ThresholdDate = (Get-Date).AddDays(-$Thresholddays)
$ArchiveFolder = Join-Path -Path $Source -ChildPath "_Archive"

#Functions-----------------------------
function Compress-Files {
    param (
        [string]$ZipFilePath,
        [string]$SourcePath
    )

    Write-Host "Processing zip - $ZipFilePath"

    # Use 7-Zip command-line interface to compress the folder with progress information
    & "$7ZipExe" a -tzip -mx9 "$ZipFilePath" "$SourcePath" | ForEach-Object {
        Write-Output $_
    }
}

#Script_Start-----------------------------

# Create the _Archive folder if it doesn't exist
if (-not (Test-Path -Path $ArchiveFolder -PathType Container)) {
    New-Item -Path $ArchiveFolder -ItemType Directory | Out-Null
}

Write-Host "Source Path is valid - $Source"

# Get files older than the threshold
Write-Host "Scanning for files - this may take a few mins"
$MatchedFiles = Get-ChildItem -LiteralPath $Source -Recurse -File | Where-Object { $_.LastWriteTime -lt $ThresholdDate }

$filecount = $MatchedFiles.Count
Write-Host "Completed - $filecount files located"

# Move matched files to the _Archive folder
Write-Host "Moving Files"
$MatchedFiles | ForEach-Object {
    $RelativePath = $_.FullName.Substring($Source.Length + 1)
    $Destination = Join-Path -Path $ArchiveFolder -ChildPath $RelativePath

    # Create the destination directory if it doesn't exist
    $DestDirectory = [System.IO.Path]::GetDirectoryName($Destination)
    if (-not (Test-Path -Path $DestDirectory -PathType Container)) {
        New-Item -Path $DestDirectory -ItemType Directory -Force | Out-Null
    }

    Move-Item -Path $_.FullName -Destination $Destination -Force
}

Write-Host "Moving log files to $ArchiveFolder Completed"

# Remove empty source folders
Write-Host "Removing empty source folders"
$MovedFolders = $MatchedFiles | ForEach-Object { $_.DirectoryName } | Select-Object -Unique
$FoldersToRemove = Get-ChildItem -LiteralPath $Source -Directory | Where-Object { $MovedFolders -contains $_.FullName }

foreach ($folder in $FoldersToRemove) {
    if (Test-Path $folder.FullName) {
        $childItems = Get-ChildItem -Path $folder.FullName -Recurse
        if ($null -eq $childItems) {
            Write-Host "Removing empty source folder $($folder.FullName)"
            Remove-Item -Path $folder.FullName -Recurse -Force
        } else {
            Write-Host "$($folder.FullName) is not empty and will not be removed."
        }
    } else {
        Write-Host "$($folder.FullName) already moved or no nested folders detected."
    }
}

# Compress Log Files or Folders in _Archive using 7zip with higher compression
$NestedFolders = Get-ChildItem -Path $ArchiveFolder -Directory

if ($NestedFolders) {
    foreach ($folder in $NestedFolders) {
        $FolderPath = $folder.FullName
        $ZipFile = "$FolderPath.zip"
        Compress-Files -ZipFilePath $ZipFile -SourcePath $FolderPath
        # Remove the folder after compression
        Remove-Item -Path $FolderPath -Recurse -Force
        Write-Host "Folder $FolderPath has been removed."
    }
} else {
    $currentDate = Get-Date -Format "yy-MM-dd"
    $ZipFile = Join-Path $ArchiveFolder "$currentDate-Archive.zip"
    Compress-Files -ZipFilePath $ZipFile -SourcePath $ArchiveFolder
    #cleanup
    $filestoRemove = Get-ChildItem -LiteralPath $ArchiveFolder -File | Where-Object { $_.Extension -eq '.log' }
    $filestoRemove | Remove-Item -WhatIf
}

Write-Host "Log Files moved to $ArchiveFolder, and compressed"

# Clean up zip files older than the specified days in _Archive folder
Write-Host "House cleaning - Remove zip files older than $ArchiveDays days in _Archive folder"
$logZipFiles = Get-ChildItem -Path $ArchiveFolder -File -Filter *.zip | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$ArchiveDays) }

foreach ($zipFile in $logZipFiles) {
    Remove-Item -Path $zipFile.FullName -Force
    Write-Host "7zip file $($zipFile.FullName) has been removed."
}
#Script_End-----------------------------
