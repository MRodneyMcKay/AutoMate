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

function open-StudentWorkReportStrand {
    
    param (
        [boolean]
        #Teljesen legyen-e kitöltve az elszámolás
        $fillCompletely,
        [datetime] $yesterday = (Get-Date).AddDays(-1)
    )

    $base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\'
        $honap = $yesterday.ToString('MMMM')
        $elszamolas = "$($base)Diák elszámolás\"
        $saveAsDir= "$($elszamolas)$($yesterday.ToString('yyyy.MM')) - $honap\"
        if (-Not (Test-Path $SaveAsDir)) {
            New-Item -Path $SaveAsDir -ItemType Directory -Force
        }

    $igeny = [ordered]@{}
    $patternsu = '^F.*(' + [regex]::escape($honap) + '|' + [regex]::escape((Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month)) +').*Strand.*\.xlsx$'
    $patternsp = '^F.*(' + [regex]::escape($honap) + '|' + [regex]::escape((Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month)) +').*Strand.*(azd|ront).*xlsx$'
    $saveAsStrand= "$($saveAsDir)$($yesterday.ToString('yyyy.MM.dd.')) - STRAND.xlsx" 
    $Empty = @{ definition = ""; hours = "" }
    $default = [ordered]@{
        Úszómester = @($Empty; $Empty; $Empty)
        Csúszda = @($Empty)
        "Vízőr (pancsoló)" = @($Empty)
        "Gyógymedence" = @($Empty)
        "Öltöző női" = @($Empty)
        "Öltöző férfi" = @($Empty)
        Kisegítő = @($Empty)
        Pénztáros = @($Empty)}
    try {
        $PathSU = Find-File -Pattern $PatternSU  -Pattern2 "^((?!Front).)*$" -PromptIfNotFound $FillCompletely -PromptTitle "Strand úszómester igényei"
        $PathSP = Find-File -Pattern $PatternSP -PromptIfNotFound $FillCompletely -PromptTitle "Strand front office igényei"
        Get-Igenyek -Muszakok $Igeny -UszomesterIgenyPath $PathSU -FrontOfficeIgenyPath $PathSP
    } catch {        
        $igeny = $default
        Write-Log -Message $_.Exception.Message -Level "WARNING"
        Write-Log -Message "Alapértelmezett igények használata" -Level "WARNING"
    } finally {
        if ($igeny.Count -ne 0) {
            $strandsablon = Create-Template $igeny $fillCompletely $("Napi diákmunka elszámolás - STRAND - " + $yesterday.ToString("yyyy. MMMM dd."))
            if (-Not (Test-Path $saveAsStrand)) {
                $strandsablon.Workbook.SaveAs($saveAsStrand,[Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
                Write-Log -Message "Az elszámolás elmentve: $savaAsFurdo"
            }
            $strandsablon.Application.visible = $true
        }
        else {
            Write-Log -Message "Nem volt igény a strand elszámolásra" -Level "INFO"
            if ($strandsablon) {
                $strandsablon.Workbook.Close($false)
                $strandsablon.Application.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($strandsablon.Application) | Out-Null
            }
        }
       
    }
    #Set-ItemProperty -Path HKCU:\Software\Script -Name YesterdaysWorkingHours -value $Today.Day
}