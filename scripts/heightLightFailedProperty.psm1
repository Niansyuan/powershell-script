# TODO: $row."Line Width (nm)" should use constants instead of hard-coded strings

function Find-FailedProperty ($data, $limits, $failDataPath) {
    handleMsgColor "Finding failed property..." "DarkGreen"

    $failedDataFile = Open-ExcelPackage -Path $failDataPath
    $worksheet = $failedDataFile.Workbook.Worksheets["Data"]

    # Iterate each row of data
    foreach ($row in $data) {
        $rowIndex = $data.IndexOf($row) + 2
        $failedProperties = @()

        # Check each property against the limits
        if ([double]$row."Line Width (nm)" -lt $limits.lineWidthMin -or [double]$row."Line Width (nm)" -gt $limits.lineWidthMax) {
            $failedProperties += "Line Width (nm)"
        }
        if ([double]$row."Film Thickness (nm)" -lt $limits.filmThicknessMin -or [double]$row."Film Thickness (nm)" -gt $limits.filmThicknessMax) {
            $failedProperties += "Film Thickness (nm)"
        }
        if ([double]$row."Planarity (nm)" -gt [double]$limits.planarityMax) {
            $failedProperties += "Planarity (nm)"
        }
        if ([double]$row."Overlay Accuracy (nm)" -gt [double]$limits.overlayAccuracyMax) {
            $failedProperties += "Overlay Accuracy (nm)"
        }
        if ([double]$row."Roughness (nm)" -gt [double]$limits.roughnessMax) {
            $failedProperties += "Roughness (nm)"
        }

        # Highlight the failed properties in red
        foreach ($property in $failedProperties) {
            $columnIndex = ($worksheet.Dimension.Start.Column + ($data[0].PSObject.Properties.Name.IndexOf($property)))
            $cell = $worksheet.Cells[$rowIndex, $columnIndex]
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)
            $cell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        }
    }

    Close-ExcelPackage $failedDataFile
    handleMsgColor "Failed property found." "DarkGreen"
}
Export-ModuleMember -Function Find-FailedProperty
