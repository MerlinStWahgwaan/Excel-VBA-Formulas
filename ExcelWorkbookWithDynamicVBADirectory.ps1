# Create a new macro-enabled Excel workbook and import VBA modules and forms
# User can set VBA_FILES_DIR and OUTPUT_DIR; both default to script directory if empty

# User-configurable directories (adjust these paths before running, or leave empty for script directory)
$VBA_FILES_DIR = ""                     # Directory containing .bas, .cls, .frm files (including subdirectories); empty = script directory
$OUTPUT_DIR = ""                        # Directory where the .xlsm file will be saved; empty = script directory
$OUTPUT_BASE_NAME = "NewWorkbookWithModules"  # Base name for the output file (without extension)

# Initialize Excel COM object
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false  # Run Excel in the background
    $excel.DisplayAlerts = $false  # Suppress alerts

    # Create a new workbook
    $workbook = $excel.Workbooks.Add()

    # Determine script directory
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

    # Determine VBA files directory (default to script directory if not set)
    $vbaDir = if ([string]::IsNullOrWhiteSpace($VBA_FILES_DIR)) { $scriptDir } else { $VBA_FILES_DIR }
    Write-Host "Using VBA files directory: $vbaDir"

    # Validate VBA files directory
    if (-not (Test-Path -Path $vbaDir)) {
        throw "VBA files directory does not exist: $vbaDir"
    }

    # Get all .bas, .cls, and .frm files in the directory and subdirectories
    $vbaFiles = Get-ChildItem -Path $vbaDir -Recurse -Include *.bas, *.cls, *.frm

    # Access the VBA project
    $vbaProject = $workbook.VBProject

    # Import each VBA file
    foreach ($file in $vbaFiles) {
        try {
            $vbaProject.VBComponents.Import($file.FullName) | Out-Null
            Write-Host "Imported: $($file.Name)"
        } catch {
            Write-Host "Failed to import: $($file.Name) - $($_.Exception.Message)"
        }
    }

    # Determine output directory (default to script directory if not set)
    $outDir = if ([string]::IsNullOrWhiteSpace($OUTPUT_DIR)) { $scriptDir } else { $OUTPUT_DIR }
    Write-Host "Using output directory: $outDir"

    # Validate output directory
    if (-not (Test-Path -Path $outDir)) {
        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $outDir"
    }

    # Generate unique output file name
    $outputFile = Join-Path -Path $outDir -ChildPath "$OUTPUT_BASE_NAME.xlsm"
    $counter = 1
    while (Test-Path -Path $outputFile) {
        $outputFile = Join-Path -Path $outDir -ChildPath "$OUTPUT_BASE_NAME$counter.xlsm"
        $counter++
    }

    # Save the workbook as macro-enabled
    $workbook.SaveAs($outputFile, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    Write-Host "New workbook created at: $outputFile"
}
catch {
    Write-Host "Error: $($_.Exception.Message)"
}
finally {
    # Clean up
    if ($workbook) {
        $workbook.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}