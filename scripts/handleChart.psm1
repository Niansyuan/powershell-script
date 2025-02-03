Import-Module ImportExcel

function Create-Chart ($data, $worksheetName, $chartPath, $chartType) {
    $data | Export-Excel -Path $chartPath -WorksheetName $worksheetName -AutoSize

    # Open Excel and create chart
    $excelPackage = Open-ExcelPackage -Path $chartPath
    $workSheet = $excelPackage.Workbook.Worksheets[$worksheetName]

    # Set up chart
    $chart = $workSheet.Drawings.AddChart("Film Thickness Distribution", $chartType)
    $chart.SetPosition(1, 0, 9, 0)
    $chart.SetSize(600, 400)

    # Data range
    $startRow = 2
    $endRow = $workSheet.Dimension.End.Row
    $yRange = "Data!D$($startRow):D$($endRow)"
    $xRange = "Data!A$($startRow):A$($endRow)"
    $chart.Series.Add($yRange, $xRange)
    $chart.Title.Text = "Film Thickness Distribution"

    Close-ExcelPackage $excelPackage
    Write-Output "Chart generated successfully at $chartPath"
}

Export-ModuleMember -Function Create-Chart
