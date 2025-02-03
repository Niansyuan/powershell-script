Import-Module "../utils/handleMsgColor.psm1"
Import-Module "./dataProcessing.psm1"
Import-Module "./handleChart.psm1" -DisableNameChecking
Import-Module "./heightLightFailedProperty.psm1"

# Paths and setup
$dateTime = Get-Date -Format "yyyyMMdd_HHmm"
$csvPath = "/Users/michelle.wang/Desktop/auto_handle_csv/semiconductor_measurements.csv"
$distPath = "/Users/michelle.wang/Desktop/auto_handle_csv/dist"
$exportPath = $distPath + "/export_" + $dateTime + ".xlsx"
$failDataPath = $distPath + "/fail_" + $dateTime + ".xlsx"

$worksheetName = "Data"
$chartType = "BarClustered"

# Quality limits
$limits = @{
    lineWidthMin = 45
    lineWidthMax = 55
    filmThicknessMin = 100
    filmThicknessMax = 150
    planarityMax = 5
    overlayAccuracyMax = 10
    roughnessMax = 2
}

# Read CSV data
$originalData = Import-Csv -Path $csvPath

# Process CSV
$data = Convert-CsvData $originalData $limits

if ($data.Count -gt 0) {
    try {
        handleMsgColor "Removing last time dist folder files." "DarkGreen"
        if (Test-Path $distPath) {
            Remove-Item -Path (Join-Path $distPath "*") -Recurse -Force
        } else {
            handleMsgColor "Creating dist folder." "DarkGreen"
            New-Item -ItemType Directory -Path $distPath | Out-Null
        }

        handleMsgColor "Processing and chart creation starting." "DarkBlue"

        Create-Chart $data $worksheetName $exportPath $chartType

        handleMsgColor "Processed data and chart successfully created. $exportPath" "DarkBlue"
    } catch {
        handleMsgColor "Error: $_" "Red"
    }
} else {
    handleMsgColor "No data found to process." "DarkYellow"
}

# Filter failed data and export separately
$failedData = $data | Where-Object { $_.Quality -eq "Fail" }

if ($failedData.Count -gt 0) {
    handleMsgColor "Exporting failedData to $failDataPath" "DarkGreen"
    $failedData | Export-Excel $failDataPath -worksheetName $worksheetName -AutoSize
    # highlight failed data by property
    Find-FailedProperty $failedData $limits $failDataPath
} else {
    handleMsgColor "No failed data found." "DarkYellow"
}