param (
    [string]$pdfPath,
    [string]$Ranges,
    [int]$PagesPerGroup
)

# Check if Python is installed
$python = Get-Command py -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "Python is not installed or not in PATH."
    exit 1
}

$scriptlocation = "$PSScriptRoot\Scripts\pdf_split\pdf_split.py"
# Check if the PDF file exists
if (-not (Test-Path $pdfPath)) {
    Write-Host "The file '$pdfPath' does not exist."
    exit 1
}

# Run the Python script with the provided arguments
try {
    Write-Host "Running the Python script to split the PDF..."
    if ($Ranges) {
        $result = py $scriptlocation $pdfPath --ranges $Ranges
        Write-Host $result

    }
    if ($PagesPerGroup) {
        $result = py $scriptlocation $pdfPath --division $PagesPerGroup
        Write-Host $result

    }
}

catch {
    Write-Host "Error running the Python script: $_"
    exit 1
}

Write-Host "PDF split process completed."
