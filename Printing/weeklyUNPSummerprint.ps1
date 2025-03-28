<#  
    This file is part of sturdy-octo-garbanzo.  

    sturdy-octo-garbanzo is free software: you can redistribute it and/or modify  
    it under the terms of the GNU General Public License as published by  
    the Free Software Foundation, either version 3 of the License, or  
    (at your option) any later version.  

    This program is distributed in the hope that it will be useful,  
    but WITHOUT ANY WARRANTY; without even the implied warranty of  
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the  
    GNU General Public License for more details.  

    You should have received a copy of the GNU General Public License  
    along with this program. If not, see <https://www.gnu.org/licenses/>.  
#>

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