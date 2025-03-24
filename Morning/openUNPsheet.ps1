Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\Printing\TestSchoolday.psm1')
# Get the current day of the week
$today = (Get-Date)

# Check if today is Saturday or Sunday
if ($today.DayOfWeek -in @('Saturday', 'Sunday')) {
    exit
}

if (Test-Schoolday) {
    $nap = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetDayName($today.DayOfWeek))
    $base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Úszó Nemzet Program\'
    $SaveAsDir="$($base)$($today.ToString('yyyy.MM')) - $((Get-Culture).DateTimeFormat.GetMonthName($today.Month))\"
    $SaveAs="$($SaveAsDir)$($today.ToString('yyyy.MM.dd')).docx"
    $Filename = "$($base)-=SABLON=-\$nap.docx"
    $Word=NEW-Object –comobject Word.Application
    $Word.visible=$true
    if (-Not ([System.IO.Directory]::Exists($SaveAsDir))) {
        mkdir -p $SaveAsDir
    }
    if (-Not ([System.IO.File]::Exists($SaveAs))) {
        $Document=$Word.documents.open($Filename)
        $Document.SaveAs2($SaveAs)
    } else{
        $Document=$Word.documents.open($SaveAs)
    }
}