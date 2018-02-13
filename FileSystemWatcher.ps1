param(
    [string]$folder, #Folder to watch
    [string]$filter, #Files to watch for eg *.zip
    [string]$title, #Title for progress bar
    [Int]$progressId, #Progress bar ID
    [string]$outPutFolder, #Folder to place unzipped files
    [string]$dtexec, #Path to dtexec
    [string]$dtsx #Path to package file
)

$fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{
    IncludeSubdirectories = $true              
    NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'
}

$pso = new-object psobject -property @{
    title = $title; 
    bar = $progressId;
    outPutFolder = $outPutFolder;
    dtexec = $dtexec;
    dtsx = $dtsx;
}

Register-ObjectEvent $fsw Created -SourceIdentifier $title -MessageData $pso -Action {                 

    Write-Progress -id 1 -Activity "SSIS FSW" -Status "----------------------------------------"
    
    $path = $Event.SourceEventArgs.FullPath
    $title = $Event.MessageData.title
    $bar = $Event.MessageData.bar
    $outPutFolder = $Event.MessageData.outPutFolder;
    $dtexec = $Event.MessageData.dtexec;
    $dtsx = $Event.MessageData.dtsx;    
    
    $name = [io.path]::GetFileNameWithoutExtension($Event.SourceEventArgs.Name)

    Write-Progress -Id $bar -ParentId 1 -Activity $title -Status $name' Downloading...' 

    While ($True)
    {
      Try { 
            [IO.File]::OpenWrite($path).Close() 
            Break
          }
      Catch { Start-Sleep -Seconds 1 }
    }  
   
    Write-Progress -Id $bar -ParentId 1 -Activity $title -Status $name' Unzipping....'     

    $zipCommand = ".\7za.exe"
    $zipArgs =  "x " + $path + " -aoa -o`"" + $outPutFolder + "\" + $name + "`""
        
    Start-Process $zipCommand -ArgumentList $zipArgs -WindowStyle Hidden | Out-Null
    Write-Progress -Id $bar -ParentId 1 -Activity $title -Status $name' Done Unzipping'   
    
    $etlArgs = "/File `"$dtsx`" /set \Package.Variables[`$Package::SourceFolder];`"$outPutFolder\$name`""

    Write-Progress -Id $bar -ParentId 1 -Activity $title -Status $name' Running ETL...'    
    Start-Process $dtexec -ArgumentList $etlArgs -WindowStyle Hidden | Out-Null
    Write-Progress -Id $bar -ParentId 1 -Activity $title -Status $name' Done'
}
