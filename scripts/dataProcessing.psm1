# TODO: $row."Line Width (nm)" should use constants instead of hard-coded strings

function Convert-CsvData($data, $limits) {
    foreach ($row in $data) {
        if ($row."Line Width (nm)" -ge $limits.lineWidthMin -and $row."Line Width (nm)" -le $limits.lineWidthMax -and
            $row."Film Thickness (nm)" -ge $limits.filmThicknessMin -and $row."Film Thickness (nm)" -le $limits.filmThicknessMax -and
            $row."Planarity (nm)" -le $limits.planarityMax -and
            $row."Overlay Accuracy (nm)" -le $limits.overlayAccuracyMax -and
            $row."Roughness (nm)" -le $limits.roughnessMax
        ) {
            $row | Add-Member -MemberType NoteProperty -Name "Quality" -Value "Pass" -Force
        } else {
            $row | Add-Member -MemberType NoteProperty -Name "Quality" -Value "Fail" -Force
        }
    }
    $data | Sort-Object -Property "Manufacturing Date"
}
Export-ModuleMember -Function Convert-CsvData
