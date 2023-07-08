
$Downloads_Path = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path

$Watcher = New-Object System.IO.FileSystemWatcher
$Watcher.Path = $Downloads_Path
$Watcher.Filter = "*.*"
$Watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName
$Watcher.IncludeSubdirectories = $false
$Watcher.EnableRaisingEvents = $true

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

$Action = {
    $File_Name = [System.Management.Automation.WildcardPattern]::Escape($Event.SourceEventArgs.Name)
    $Origin_File_Full_Path = Join-Path -Path $Downloads_Path -ChildPath $File_Name
    $extension = [System.IO.Path]::GetExtension($Origin_File_Full_Path)

    $Contains_Extension = $File_Extensions.ContainsKey($extension)
    $Type_Folder = if ($Contains_Extension -eq $true) { $File_Extensions[$extension] } else { "Other" }
    $Destination_Path = if (-not($Type_Folder -eq "Torrents")) { Join-Path -Path $Downloads_Path -ChildPath $Type_Folder } else { "F:\Torrents" }    

    if (-not (Test-Path -Path $Destination_Path)) { New-Item -ItemType Directory -Path $Destination_Path | Out-Null }

    $Destination_File_Full_Path = Join-Path -Path $Destination_Path -ChildPath $File_Name
    
    $Iterator = 2
    while (Test-Path -Path $Destination_File_Full_Path) {
        Write-Host "# [$File_Name] File already exists in the destination folder."
        # File already exists in the destination folder, rename the file
        $New_File_Name = [System.IO.Path]::GetFileNameWithoutExtension($File_Name) + " ($Iterator)$extension"
        $Destination_File_Full_Path = Join-Path -Path $Destination_Path -ChildPath $New_File_Name

        $Iterator = $Iterator + 1
    }

    do { 
         try { Move-Item -Path $Origin_File_Full_Path -Destination $Destination_File_Full_Path -Force }
         catch  { Write-Host "An error occurred:" ; Write-Host $_ ; Start-Sleep -Seconds 1 }
    } while (-not (Test-Path -Path $Destination_File_Full_Path))

    Write-Host "# [$File_Name] File moved successfully from '$Origin_File_Full_Path' to '$Destination_File_Full_Path'`n"
}

Register-ObjectEvent -InputObject $Watcher -EventName Created -Action $Action

while ($true) {
    Start-Sleep -Seconds 1
}