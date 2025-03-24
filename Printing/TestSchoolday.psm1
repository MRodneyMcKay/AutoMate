function Test-Schoolday {
    param (
        [datetime]$date = (Get-Date)
    )

    $schooldayRanges = @(
        [PSCustomObject]@{ Start = '2024-10-26'; End = '2024-11-03' },
        [PSCustomObject]@{ Start = '2024-12-21'; End = '2025-01-05' },
        [PSCustomObject]@{ Start = '2025-04-17'; End = '2025-04-27' },
        [PSCustomObject]@{ Start = '2025-06-21'; End = '2025-09-01' }
    )

    foreach ($range in $schooldayRanges) {
        if ($date -ge (Get-Date $range.Start) -and $date -le (Get-Date $range.End)) {
            return $false
        }
    }

    return $true
}
