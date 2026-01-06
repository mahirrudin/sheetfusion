# XLS to XLSX Converter for Windows using Microsoft Excel
# Usage: .\convert_xls_to_xlsx.ps1 -Input "input.xls" [-Output "output.xlsx"]
#        .\convert_xls_to_xlsx.ps1 -Input "C:\path\to\directory"

param(
    [Parameter(Mandatory=$true, HelpMessage="Input .xls file or directory")]
    [string]$Input,
    
    [Parameter(Mandatory=$false, HelpMessage="Output .xlsx file (optional)")]
    [string]$Output = ""
)

# Excel file format constants
$xlOpenXMLWorkbook = 51  # .xlsx format

# Function to check if Excel is installed
function Test-ExcelInstalled {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

# Function to convert a single file
function Convert-XlsToXlsx {
    param(
        [string]$InputFile,
        [string]$OutputFile = ""
    )
    
    # Get absolute paths
    $InputPath = (Resolve-Path $InputFile).Path
    
    if ($OutputFile -eq "") {
        # Auto-generate output filename
        $OutputPath = [System.IO.Path]::ChangeExtension($InputPath, ".xlsx")
    }
    else {
        $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
    }
    
    Write-Host "Converting: $InputFile" -ForegroundColor Yellow
    
    try {
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Open the .xls file
        $workbook = $excel.Workbooks.Open($InputPath)
        
        # Save as .xlsx
        $workbook.SaveAs($OutputPath, $xlOpenXMLWorkbook)
        
        # Close workbook
        $workbook.Close($false)
        
        # Quit Excel
        $excel.Quit()
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "✓ Converted: $(Split-Path $OutputPath -Leaf)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Error converting file: $_" -ForegroundColor Red
        
        # Cleanup on error
        if ($workbook) {
            $workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        
        return $false
    }
}

# Main script
Write-Host ""
Write-Host "XLS to XLSX Converter (Windows)" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Check if Excel is installed
if (-not (Test-ExcelInstalled)) {
    Write-Host "Error: Microsoft Excel is not installed." -ForegroundColor Red
    Write-Host "Please install Microsoft Office to use this script." -ForegroundColor Red
    exit 1
}

# Check if input exists
if (-not (Test-Path $Input)) {
    Write-Host "Error: File or directory not found: $Input" -ForegroundColor Red
    exit 1
}

# Check if input is a directory
if (Test-Path $Input -PathType Container) {
    Write-Host "Converting all .xls files in: $Input" -ForegroundColor Yellow
    Write-Host ""
    
    # Get all .xls files (not .xlsx)
    $xlsFiles = Get-ChildItem -Path $Input -Filter "*.xls" -File | Where-Object { $_.Extension -eq ".xls" }
    
    if ($xlsFiles.Count -eq 0) {
        Write-Host "No .xls files found in directory." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "Found $($xlsFiles.Count) .xls file(s) to convert" -ForegroundColor Cyan
    Write-Host ""
    
    $successCount = 0
    $failedCount = 0
    
    foreach ($file in $xlsFiles) {
        if (Convert-XlsToXlsx -InputFile $file.FullName) {
            $successCount++
        }
        else {
            $failedCount++
        }
    }
    
    Write-Host ""
    Write-Host "✓ Successfully converted $successCount of $($xlsFiles.Count) file(s)" -ForegroundColor Green
    if ($failedCount -gt 0) {
        Write-Host "✗ Failed to convert $failedCount file(s)" -ForegroundColor Red
    }
}
else {
    # Single file conversion
    if (-not $Input.EndsWith(".xls")) {
        Write-Host "Error: Input file must have .xls extension" -ForegroundColor Red
        exit 1
    }
    
    if (Convert-XlsToXlsx -InputFile $Input -OutputFile $Output) {
        Write-Host ""
        Write-Host "✓ Conversion complete!" -ForegroundColor Green
    }
    else {
        exit 1
    }
}

Write-Host ""
