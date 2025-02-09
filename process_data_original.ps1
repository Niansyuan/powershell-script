# Import files
Import-Module "./utils/handleMsgColor.psm1"

# get date time
$dateTime = Get-Date -Format "yyyyMMdd_HHmm"

# change the path by the file path in your computer
$csvPath = "/Users/michelle.wang/Desktop/auto_handle_csv/semiconductor_measurements.csv"
$distPath = "/Users/michelle.wang/Desktop/auto_handle_csv/dist"

$processData_CSV = $distPath + "/process_" + $dateTime + ".csv"
$processData_HTML = $distPath + "/process_" + $dateTime + ".html"
$failData_CSV = $distPath + "/fail_" + $dateTime + ".csv"

# Read CSV data
$data = Import-Csv -Path $csvPath

# Define quality criteria
$lineWidthMin = 45
$lineWidthMax = 55
$filmThicknessMin = 100
$filmThicknessMax = 150
$planarityMax = 5
$overlayAccuracyMax = 10
$roughnessMax = 2

# Process each row and determine pass/fail
foreach ($row in $data) {
    if ($row."Line Width (nm)" -ge $lineWidthMin -and $row."Line Width (nm)" -le $lineWidthMax -and
        $row."Film Thickness (nm)" -ge $filmThicknessMin -and $row."Film Thickness (nm)" -le $filmThicknessMax -and
        $row."Planarity (nm)" -le $planarityMax -and
        $row."Overlay Accuracy (nm)" -le $overlayAccuracyMax -and
        $row."Roughness (nm)" -le $roughnessMax
    ) {
        $row | Add-Member -MemberType NoteProperty -Name "Quality" -Value "Pass" -Force
    } else {
        $row | Add-Member -MemberType NoteProperty -Name "Quality" -Value "Fail" -Force
    }
}

# Display processed data by condition
$data = $data | Sort-Object -Property "Manufacturing Date"

# Export processed results
if($data.Count -gt 0) {
    try {
        Write-Output "Removing last time dist folder files."
        if (Test-Path $distPath) {
            Remove-Item -Path (Join-Path $distPath "*") -Recurse -Force
        } else {
            Write-Output "Creating dist folder."
            New-Item -ItemType Directory -Path $distPath | Out-Null
        }

        handleMsgColor "Exporting CSV to $processData_CSV" "DarkGreen"
        # Export CSV
        $data | Export-Csv $processData_CSV -NoTypeInformation
        # Export HTML
        $data | ConvertTo-Html | Out-File $processData_HTML
    } catch {
        handleMsgColor "Error: $_" "Red"
    }
} else {
    handleMsgColor "No data found." "DarkYellow"
}

# Filter failed data and export separately
$failedData = $data | Where-Object { $_.Quality -eq "Fail" }

if($failedData.Count -gt 0) {
    handleMsgColor "Exporting failedData to $failData_CSV" "DarkGreen"
    $failedData | Export-Csv $failData_CSV -NoTypeInformation
}else {
    handleMsgColor "No failed data found." "DarkYellow"
}

handleMsgColor "Processed data saved successfully. Please check $distPath" "DarkBlue"

# Install required module
Import-Module ImportExcel

# Prepare Excel file path
$chartPath = $distPath + "/DataChart_" + $dateTime + ".xlsx"

# Create a simple Excel file with the processed data
$data | Export-Excel -Path $chartPath -WorksheetName "Data"

# Create a histogram chart for Film Thickness distribution
$excelPackage = Open-ExcelPackage -Path $chartPath
$ws = $excelPackage.Workbook.Worksheets["Data"]

# Set up the chart
$chart = $ws.Drawings.AddChart("Film Thickness Distribution", "BarClustered")
$chart.SetPosition(1, 0, 9, 0) # Row, offset, column, offset
$chart.SetSize(600, 400)

# Determine row count from the worksheet itself
$startRow = 2
$endRow = $ws.Dimension.End.Row

# Define X and Y data ranges (Manufacturing date is in column A and film thickness in column D)
$yRange = "Data!D$($startRow):D$($endRow)"
$xRange = "Data!A$($startRow):A$($endRow)"
$chart.Series.Add($yRange, $xRange)
$chart.Title.Text = "Film Thickness Distribution"

# Save and close the package
Close-ExcelPackage $excelPackage

Write-Output "Chart generated successfully at $chartPath"