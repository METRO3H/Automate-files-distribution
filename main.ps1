<# Import-Module .\getDestinationPath.ps1 #>

$downloadsPath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $downloadsPath
$watcher.Filter = "*.*"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true


$action = {
    $file = $Event.SourceEventArgs.Name
    $fileFullPath = Join-Path -Path $downloadsPath -ChildPath $file
    $changeType = $Event.SourceEventArgs.ChangeType
    # Check if the event is for a file creation
    if (-not($changeType -eq "Created") -or -not(Test-Path -Path $fileFullPath -PathType Leaf)) {
        Write-Host "The triggered event is not caused by a file or the event is not 'created' type"
        Write-Host "Event Type : " $changeType
        Write-Host "Path : " $fileFullPath
        return
    }
    $extension = [System.IO.Path]::GetExtension($fileFullPath)

    $File_Extensions = @{
        #Videos
        ".mp4" = "Videos" ; ".mkv" = "Videos"
        #Music
        ".mp3" = "Music" ; ".wav" = "Music"; ".flac" = "Music"; ".m4a" = "Music"   
        #Images
        ".png" = "Images" ; ".webp" = "Images"; ".jpg" = "Images"; ".jpeg" = "Images"            
        #Documents
        ".txt" = "Documents\Text" ; ".pdf" = "Documents\PDF"; ".docx" = "Documents\Word"; ".pptx" = "Documents\Power point";
        ".xlsx" = "Documents\Excel"
        #Compressed
        ".zip" = "Compressed"; ".rar" = "Compressed" ; ".tar" = "Compressed"; ".7z" = "Compressed"
        #Setups
        ".exe" = "Setups"
        #Torrents
        ".torrent" = "Torrents"                    
    }

    $Type_Folder = ""

    if ($File_Extensions.ContainsKey($extension)) {
        Write-Host "Case extension $extension found."
        $Type_Folder = $File_Extensions[$extension]

<#          if($Type_Folder == "Torrents"){
           return "F:\Torrents"
       } #>
    }
    else {
        Write-Host "Case not found."
        $Type_Folder = "Other"
    }

    $destinationPath = Join-Path -Path "C:\bob" -ChildPath $Type_Folder

    if (-not (Test-Path -Path $destinationPath)) {
        New-Item -ItemType Directory -Path $destinationPath | Out-Null
    }

    $destinationFileFullPath = Join-Path -Path $destinationPath -ChildPath $file
    $iterator = 2
    while (Test-Path -Path $destinationFileFullPath) {
        Write-Host "File already exists in the destination folder: $file"
        # File already exists in the destination folder, rename the file
        $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($file) + " ($iterator)$extension"
        $destinationFileFullPath = Join-Path -Path $destinationPath -ChildPath $newFileName

        $iterator = $iterator + 1
    }

    do {
                 
        try {
            Move-Item -Path $fileFullPath -Destination $destinationFileFullPath -Force
        }
        catch [System.IO.IOException] {
            Write-Host "Sharing violation occurred while moving the file: $file"
            Write-Host "Retrying after a short delay..."
            Start-Sleep -Seconds 1
        }


    } while (-not (Test-Path -Path $destinationFileFullPath))

}

Register-ObjectEvent -InputObject $watcher -EventName Created -Action $action

while ($true) {
    Start-Sleep -Seconds 1
}

