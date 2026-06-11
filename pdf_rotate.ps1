<#
.SYNOPSIS
Rotates a PDF file using a Python script (rotate_pdf.py).

.DESCRIPTION
This script acts as a wrapper around a Python script that uses the PyPDF2 library
to rotate all pages in a PDF file. It passes user-specified arguments such as
rotation angle, direction, and output file name.

.PARAMETER InputFile
Path to the input PDF file.

.PARAMETER Rotation
Rotation angle in degrees. Valid values are:
- 90
- 180
- 270

.PARAMETER Direction
Direction of rotation:
- cw  = clockwise
- ccw = counter-clockwise

.PARAMETER OutputFile
(Optional) Path to save the rotated PDF.
If not provided, a new file is created with "_rotated" appended to the name.
You can also specify the same name as the input file to overwrite it.

.EXAMPLES

# Rotate 90 degrees clockwise and create a new file
.\rotate_pdf.ps1 -InputFile "file.pdf" -Rotation 90 -Direction cw

# Rotate 180 degrees counter-clockwise and specify output name
.\rotate_pdf.ps1 -InputFile "file.pdf" -Rotation 180 -Direction ccw -OutputFile "output.pdf"

# Rotate and overwrite original file
.\rotate_pdf.ps1 -InputFile "file.pdf" -Rotation 90 -Direction cw -OutputFile "file.pdf"

.NOTES
Requirements:
- Python must be installed and available in PATH
- PyPDF2 must be installed (pip install PyPDF2)
- rotate_pdf.py must exist in the same directory or update the path below

#>


param (
    [Parameter(Mandatory=$true)]
    [string]$InputFile,

    [Parameter(Mandatory=$true)]
    [ValidateSet(90,180,270)]
    [int]$Rotation,

    [Parameter(Mandatory=$true)]
    [ValidateSet("cw","ccw")]
    [string]$Direction,

    [string]$OutputFile
)

# Path to the Python script
# Modify this if your script is in a different locatio

$PythonScript = "$PSScriptRoot\Scripts\pdf_rotate\pdf_rotate.py"

# Validate input file exists
if (!(Test-Path $InputFile)) {
    Write-Error "Input file not found: $InputFile"
    exit 1
}

# Construct arguments
$argsList = @($PythonScript, $InputFile, $Rotation, $Direction)

if ($OutputFile) {
    $argsList += $OutputFile
}


# Execute Python script
Write-Host "Running PDF rotation..."
py @argsList


# Inform user of completion
if ($LASTEXITCODE -eq 0) {
    Write-Host "Rotation completed successfully." -ForegroundColor Green
} else {
    Write-Error "Rotation failed."
}

