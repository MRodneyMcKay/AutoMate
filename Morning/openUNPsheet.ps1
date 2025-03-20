# Get the current day of the week
$today = (Get-Date).DayOfWeek

# Check if today is Saturday or Sunday
if ($today -eq 'Saturday' -or $today -eq 'Sunday') {
    Exit
}

if ($today -ge (Get-Date "2024-10-26") -and $today -le (Get-Date "2024-11-03") -or
        $today -ge (Get-Date "2024-12-21") -and $today -le (Get-Date "2025-01-05") -or
        $today -ge (Get-Date "2025-04-17") -and $today -le (Get-Date "2025-04-27") -or
        $today -eq (Get-Date "2025-05-01")) {
        Write-Output "Holiday detected. Skipping printing."
    }else{
        $nap = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetDayName($today.DayOfWeek))
        $base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Úszó Nemzet Program\'
        $SaveAsDir="$($base)$($today.ToString('yyyy.MM')) - $((Get-Culture).DateTimeFormat.GetMonthName($today.Month))\"
        $SaveAs="$($SaveAsDir)$($today.ToString('yyyy.MM.dd')).docx"
        $SaveAs
        $Filename = "$($base)-=SABLON=-\$nap.docx"
        $Word=NEW-Object –comobject Word.Application
        $Word.visible=$true
        if (-Not ([System.IO.Directory]::Exists($SaveAsDir)))
        {
            mkdir -p $SaveAsDir
        }
        if (-Not ([System.IO.File]::Exists($SaveAs)))
        {
            $Document=$Word.documents.open($Filename)
            $Document.SaveAs2($SaveAs)
        }
        else
        {
            $Document=$Word.documents.open($SaveAs)
        }
    }

