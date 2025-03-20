$today = Get-Date
$honap = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName((get-date).Month))
$base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Úszó Nemzet Program\'
$Filename = "$($base)-=SABLON=-\nyár.docx"
$SaveAsDir="$($base)$($today.Year).$($today.Month) - $honap\"
$SaveAs="$($SaveAsDir)$($today.ToString('yyyy.MM.dd.')).docx"
$Word=NEW-Object –comobject Word.Application
Set-ItemProperty -Path $Filename -Name IsReadOnly -Value $true
$Document=$Word.documents.open($Filename)
$FindText = "Nap"
$ReplaceText = $today.ToString("MMMM dd.") # <= Replace it with this text
$MatchCase = $false
$MatchWholeWorld = $true
$MatchWildcards = $false
$MatchSoundsLike = $false
$MatchAllWordForms = $false
$Forward = $false
$Wrap = 1
$Format = $false
$Replace = 2
if (!$($Document.Content.Find.Execute($FindText, $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $ReplaceText, $Replace))) {
    $Document.Close(-1)
    $Word.Quit()
}
else {   

    if (-Not ([System.IO.Directory]::Exists($SaveAsDir)))
    {
        mkdir -p $SaveAsDir
    }
    $Document.SaveAs2($SaveAs)
    $Word.visible=$true
}