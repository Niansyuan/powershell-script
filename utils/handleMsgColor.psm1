function handleMsgColor {
    param (
        [string]$text,
        [string]$color
    )
    Write-Host $text -ForegroundColor $color
}

Export-ModuleMember -Function handleMsgColor