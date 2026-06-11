# Specify the path to your Excel file
$excelFilePath = $PSScriptRoot + "\Scripts\morn_pull\data_entry.xlsx"
 
Write-Host "Opening data_entry.xlsx for pasting of aspen data in location"
Write-Host $excelFilePath

$Global:finishScript = $true
$Global:lastEventTime = "empty"

if (Test-Path -Path $excelFilePath) {
    # File exists, so open it
    $ExcelObj = New-Object -ComObject Excel.Application
    $null = $ExcelObj.Workbooks.Open(
        $excelFilePath, #File name 
        $null, # Updatelinks
        $false, #ReadOnly
        [type]::Missing, #Format
        [type]::Missing, #Password
        [type]::Missing, #WriteResPassword
        $true  #IgnoreReadonlyRecommended
        )
    
    $ExcelObj.Visible = $true
    Write-Host "Excel file opened."
} else {
    # File doesn't exist, create a new one
    $ExcelObj = New-Object -ComObject Excel.Application
    $null = $ExcelObj.Workbooks.Add()
    $null = $ExcelObj.Workbooks.Open(
        $excelFilePath, #File name 
        $null, # Updatelinks
        $false, #ReadOnly
        [type]::Missing, #Format
        [type]::Missing, #Password
        [type]::Missing, #WriteResPassword
        $true  #IgnoreReadonlyRecommended
        )
    $ExcelObj.Visible = $true
    Write-Host "New Excel file created."
}

# Create a FileSystemWatchers to monitor changes in the folder
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = (Get-Item $excelFilePath).DirectoryName
$watcher.Filter = "*data_entry.xlsx"
$watcher.IncludeSubdirectories = $false

$watcher2 = New-Object System.IO.FileSystemWatcher
$watcher2.Path = (Get-Item $excelFilePath).DirectoryName
$watcher2.Filter = "*~`$data*.xlsx"
$watcher2.IncludeSubdirectories = $false

# Define the action to take when the file is saved
$action = {
    $currentTimeStamp = $Event.TimeGenerated.ToString()
    # Write-Host $Global:lastEventTime
    # Write-Host $currentTimeStamp

    if ($Global:lastEventTime -ne $currentTimeStamp) {

        Write-Host "Data Entry saved! Executing Python script..."
        # Replace with the path to your Python script
        $pythonScriptPath = $PSScriptRoot + "\Scripts\morn_pull\morn_pull.py"
        Start-Process py -ArgumentList "$pythonScriptPath $excelFilePath" -NoNewWindow
    }
    $Global:lastEventTime = $currentTimeStamp
}

# Define the action to take when the file is deleted or renamed
$deleteAction = {
    Write-Host "Data Entry closed. Stopping monitoring..."
    # Clean up and stop monitoring
    if ($Global:finishScript -eq $true) {
        $Global:finishScript = $false
    }
}


# Register the event handlers
Register-ObjectEvent -InputObject $watcher "Renamed" -SourceIdentifier "Saved_file" -Action $action
Register-ObjectEvent -InputObject $watcher2 "Deleted" -SourceIdentifier "Delete_End" -Action $deleteAction

# Start monitoring
$watcher.EnableRaisingEvents = $true
$watcher2.EnableRaisingEvents = $true

# Keep the script running
try{
    while ($Global:finishScript) {
        Start-Sleep -Seconds 1 #adjust as required
    }
}
finally {
    <#Do this after the try block regardless of whether an exception occurred or not#>
    # Clean up when the script is terminated
    Write-Host "finishing script"
    Unregister-Event -SourceIdentifier "Saved_file"
    Unregister-Event -SourceIdentifier "Delete_End"
    
    $watcher.Dispose()
    $watcher2.Dispose()
    exit
}
