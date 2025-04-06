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

function open-StudentWorkReportFurdo {
    
    param (
        [boolean]  $fillCompletely = $false,
        [datetime] $yesterday = (Get-Date).AddDays(-1)
    )
    $base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\'
    $honap = $yesterday.ToString('MMMM')
    $elszamolas = "$($base)Diák elszámolás\"
    $saveAsDir= "$($elszamolas)$($yesterday.ToString('yyyy.MM')) - $honap\"
    if (-Not (Test-Path $SaveAsDir)) {
        New-Item -Path $SaveAsDir -ItemType Directory -Force
    }
    $savaAsFurdo= "$($saveAsDir)$($yesterday.ToString('yyyy.MM.dd.')).xlsx"

    $igeny = [ordered]@{}

    $patternfu = '^F.*(' + [regex]::escape($honap) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month) +').*.*ürdő.*\.xlsx$'
    $patternfp =  '^F.*(' + [regex]::escape($honap) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month) +').*Fürdő.*(azd|ront).*xlsx$'
    $Empty = @{ definition = ""; hours = "" }
    $default = [ordered] @{"Úszómester" = @($empty;$empty)
    "Csúszda" = @($empty)
    "Gyógymedence`nfelügyelet" = @($empty;$empty)
    "Öltöző szolgálat" = @($empty;$empty)
    Gyermekmegőrző = @($empty)
    "Pénztáros" = @($empty;$empty)
    Recepció = @($empty)}
    try {
        $PathFU = Find-File -Pattern $PatternFU -Pattern2 "^((?!Front).)*$" -PromptIfNotFound $FillCompletely -PromptTitle "Fürdő úszómester igényei"
        $PathFP = Find-File -Pattern $patternfp -PromptIfNotFound $FillCompletely -PromptTitle "Fürdő front office igényei"
        Get-Igenyek -Muszakok $Igeny -UszomesterIgenyPath $PathFU -FrontOfficeIgenyPath $PathFP
    } catch {        
        $igeny = $default
        Write-Log -Message $_.Exception.Message -Level "WARNING"
        Write-Log -Message "Alapértelmezett igények használata" -Level "WARNING"
    } finally {
        $furdosablon = Create-Template $igeny $fillCompletely $("Napi diákmunka elszámolás " + $yesterday.ToString("yyyy. MMMM dd."))
        if (-Not (Test-Path $savaAsFurdo)) {
            $furdosablon.Workbook.SaveAs($savaAsFurdo,[Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
            Write-Log -Message "Az elszámolás elmentve: $savaAsFurdo"
        }
        $furdosablon.Application.visible=$true
    }
    Set-ItemProperty -Path HKCU:\Software\Script -Name YesterdaysWorkingHours -value $Today.Day
}