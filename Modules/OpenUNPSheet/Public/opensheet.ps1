<#  
    This file is part of AutoMate.  

    AutoMate is free software: you can redistribute it and/or modify  
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

function open-UNP {
    param (
        $today = (Get-Date)
    )

    if (Test-Schoolday) {
        $nap = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetDayName($today.DayOfWeek))
        $base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Úszó Nemzet Program\'
        $SaveAsDir="$($base)$($today.ToString('yyyy.MM')) - $((Get-Culture).DateTimeFormat.GetMonthName($today.Month))\"
        $SaveAs="$($SaveAsDir)$($today.ToString('yyyy.MM.dd')).docx"
        $Filename = "$($base)-=SABLON=-\$nap.docx"
        $Word=NEW-Object –comobject Word.Application
        $Word.visible=$true
        if (-Not (Test-Path $SaveAsDir)) {
            New-Item -Path $SaveAsDir -ItemType Directory -Force
            Write-Log -Message "Directory created: $SaveAsDir"
        }
        if (-Not (Test-Path $SaveAs)) {
            try {
                $Document=$Word.documents.open($Filename)
                if ($Document) {
                    Write-Log -Message  "$Filename opened successfully."
                    $Document.SaveAs2($SaveAs)
                    Write-Log -Message "Document created: $SaveAs"
                }
                else {
                    throw
                }               
            }
            catch {
                Write-Log -Message "Error opening document: $Filename" -Level "ERROR"
            }
            
        } else{
            $Document=$Word.documents.open($SaveAs)
            Write-Log -Message "Document opened: $SaveAs"
        }
    }
    else {
        Write-Log -Message "Iskola szünet van"
    }
}

function open-SummerCamp {
    param (
        $today = (Get-Date)
    )
    Write-Log -Message "Testing if $($today.ToString('yyyy-MM-dd')) is a summer day." -Level "INFO"
    if (-not (Test-Summercamp)) {
        Write-Log -Message "Not a summer camp day."
        return
    }
    $honap = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName($today.Month))
    $base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Úszó Nemzet Program\'
    $Filename = "$($base)-=SABLON=-\nyár.docx"
    $SaveAsDir="$($base)$($today.ToString('yyyy.MM')) - $honap\"
    $SaveAs="$($SaveAsDir)$($today.ToString('yyyy.MM.dd')).docx"
    $Word=NEW-Object –comobject Word.Application
    $Word.visible=$false
    $Document=$Word.documents.open($Filename)
    if (!$Document) {
        Write-Log -Message "Failed to open the document: $Filename" -Level "INFO"
        $Word.Quit()
        return
    }
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
        Write-Log -Message "Failed to replace text '$FindText' with '$ReplaceText' in the document."
    }
    else {   

        if (-Not ([System.IO.Directory]::Exists($SaveAsDir)))
        {
            mkdir -p $SaveAsDir
        }
        $Document.SaveAs2($SaveAs)
        Write-Log -Message "Document opened: $SaveAs"
        $Document.PrintOut()
        Write-Log -Message "Printed document: $Filename"
        $Word.visible=$true
    }
}