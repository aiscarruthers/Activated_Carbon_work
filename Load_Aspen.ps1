# Define the path to the Excel file
$excelFilePath = "G:\World\11 LABORATORY\7.  Analytical Results\7.6 MBR & Effluent\1 EFFLUENT ANALYSIS\Aspen_data_for_effluent_sheet.xlsm"
Write-Host "Opening File"
# Start Excel and open the file
Start-Process -FilePath "excel.exe"  `"$excelFilePath`"

Write-Host "Sleeping for 10s"

# Wait for a specified duration to allow the file to load and custom functions to run
Start-Sleep -Seconds 10

# # Get the process ID of the specific Excel instance
# $excelProcess = Get-Process -Name "EXCEL" | Where-Object { $_.MainWindowTitle -like "*Aspen_data_for_effluent_sheet.xlsm*" }

# # Close the specific Excel process
# Stop-Process -Id $excelProcess.Id -Force

# Define the window title of the Excel file you want to close
$windowTitle = "*Aspen_data_for_effluent_sheet.xlsm*"

Write-Host "finished sleep"

# Get the process ID of the specific Excel instance
$excelProcess = Get-Process -Name "Excel" | Where-Object { $_.MainWindowTitle -like $windowTitle }

Write-Host $excelProcess

# Check if the process was found
if ($excelProcess) {
    # Close the specific Excel process
    Stop-Process -Id $excelProcess.Id -Force
    Write-Output "Excel process with window title '$windowTitle' has been closed."
} else {
    Write-Output "No Excel process with window title '$windowTitle' was found."
}
