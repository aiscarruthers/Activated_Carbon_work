
Write-Host "Running Python Script to collection O2 data from the MBR"
$pythonScriptPath = $PSScriptRoot + "\Scripts\MBR_O2_rep\O2_report.py"
Start-Process python -ArgumentList "$pythonScriptPath" -NoNewWindow
