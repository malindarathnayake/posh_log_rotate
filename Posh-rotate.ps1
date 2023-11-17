#-----------------------------------------------
#-------------------- INFOS --------------------
#Script:   poshLogsRotation
#Version:  1.0.1.2
#Date:     11/17/2023
#Author:   Malinda Rathnayake

#varibles-----------------------------
param (
    [string]$Source,
    [int]$Thresholddays
)

if (-not $Source){
    Write-Output "Please provide a source path using the -Source parameter. Example:"
    Write-Output ".\posh_Rotate.ps1 -Source 'D:\your_log_Path_root'"
    return
}
$MatchedFiles = $null
$ThresholdDate = (Get-Date).AddDays(-$Thresholddays)
$ArchiveFolder = Join-Path -Path $Source -ChildPath "_Archive"
$7ZipExe = "C:\Program Files\7-Zip\7z.exe"  # Adjust the path to your 7zip executable
$filecount = $MatchedFiles.count

##Functions-----------------------------
function 7zCompress-objects {
    param (
        [string]$ZipFilePath,
        [string]$SourcePath
    )

    $7ZipExe = "C:\Program Files\7-Zip\7z.exe"

    # Compress the folder using 7zip with higher compression level (-mx9)
    Write-Output "Processing zip - $ZipFilePath"

    # Use 7-Zip command-line interface to compress the folder with progress information
    & "$7ZipExe" a -tzip -mx9 "$ZipFilePath" "$SourcePath" | ForEach-Object {
        Write-Output $_
}}

#script_Start-----------------------------

# Create the _Archive folder if it doesn't exist
if (-not (Test-Path -Path $ArchiveFolder -PathType Container)) {
    New-Item -Path $ArchiveFolder -ItemType Directory
}

Write-Output "Source Path is valid - $Source"

# Get files older than 30 days
Write-Output "Scanning for files - this may take a few mins"
$MatchedFiles = Get-ChildItem -LiteralPath $Source -Recurse | Where-Object { $_.LastWriteTime -lt $ThresholdDate }

$filecount = $MatchedFiles.count
Write-Output "Completed - $filecount files located"

# Move matched files to the _Archive folder
Write-Output "Moving Files/folders"
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

Write-Output "Moving log files to $ArchiveFolder Completed"

# Remove empty source folders
Write-Output "Removing empty source folders"

#this step will do an extra check to confirm the folder list to Remove.
$MovedFolders = $MatchedFiles | ForEach-Object { $_.DirectoryName }
$FoldersToRemove = Get-ChildItem -LiteralPath $Source -Directory | Where-Object { $MovedFolders -contains $_.FullName }

foreach ($folder in $FoldersToRemove) {
    if (Test-Path $folder.FullName) {
        if ((Get-ChildItem -Path $folder.FullName -Recurse | Measure-Object).Count -eq 0) {
            Write-Output "Removing empty source folder $($folder.FullName)"
            Remove-Item -Path $folder.FullName -Recurse -Force
        } else {
            Write-Output "$($folder.FullName) is not empty and will not be removed."
        }
    } else {
        Write-Output "$($folder.FullName) already moved or no nested folders detected."
    }
}

# Compress Log Files or Folders in _Archive using 7zip with higher compression
$NestedFolders = Get-ChildItem -Path $ArchiveFolder -Directory
if (-not $NestedFolders -eq $null){ 
  foreach ($folder in $NestedFolders) {
        $FolderPath = $folder.FullName
        $ZipFile = "$FolderPath.zip"
        7zCompress-objects -ZipFilePath $ZipFile -SourcePath $FolderPath
        # Remove the folder after compression
        Remove-Item -Path $FolderPath -Recurse -Force
        Write-Output "Folder $FolderPath has been removed."
}} else {
            $currentDate = Get-Date -Format "yy-MM-dd"
            $ZipFile = $ArchiveFolder+'\'+ "$currentDate-Archive.zip"
            7zCompress-objects -ZipFilePath $ZipFile -SourcePath $ArchiveFolder
            #cleanup
            $filestoRemove = Get-ChildItem -LiteralPath $ArchiveFolder |? { $_.Extension -eq '.log' }
            foreach ($zfile in $filestoRemove) {
            Remove-Item -Path $zfile.FullName -WhatIf
            }}

Write-Output "Log Files moved to $ArchiveFolder, and compressed"

# Clean up zip files older than 60 days in _Archive folder----------------------
Write-Output "House cleaning - Remove zip files older than 60 days in _Archive folder"
$logZipFiles = Get-ChildItem -Path $ArchiveFolder -File -Filter *.zip | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-60) }

foreach ($zipFile in $logZipFiles) {
    Remove-Item -Path $zipFile.FullName -Force
    Write-Output "7zip file $($zipFile.FullName) has been removed."
}
#script_end-----------------------------
