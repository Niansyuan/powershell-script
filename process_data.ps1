# remove last time dist folder files
Remove-Item -Path "/Users/michelle.wang/Desktop/auto_handle_csv/dist/*" -Recurse -Force

# 1. Read CSV data
$csvPath = "/Users/michelle.wang/Desktop/auto_handle_csv/semiconductor_measurements.csv"
$data = Import-Csv -Path $csvPath
$data | Get-Member

Write-Output "data" $data

# 2. Define quality criteria
$lineWidthMin = 45
$lineWidthMax = 55
$filmThicknessMin = 100
$filmThicknessMax = 150
$planarityMax = 5
$overlayAccuracyMax = 10
$roughnessMax = 2

# 3. Process each row and determine pass/fail
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

# 5. Export processed results
if($data) {
    $data | Export-Csv "/Users/michelle.wang/Desktop/auto_handle_csv/dist/processed_semiconductor_data.csv" -NoTypeInformation
}

# Filter failed data and export separately
$failedData = $data | Where-Object { $_.Quality -eq "Fail" }

if($failedData) {
    $failedData | Export-Csv "/Users/michelle.wang/Desktop/auto_handle_csv/dist/failed_semiconductor_data.csv" -NoTypeInformation
}

Write-Output "Processed data saved successfully."
