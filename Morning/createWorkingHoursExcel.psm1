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

Import-Module "$PSScriptRoot\createTemplate.psm1"
Add-Type -AssemblyName PresentationFramework
[System.Reflection.Assembly]::LoadFrom("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll")
[System.Reflection.Assembly]::LoadFrom("C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll ") 

function Show-ErrorMessage {
    param (
        [string]$Message,
        [string]$Title = "Hiba"
    )
    [System.Windows.Forms.MessageBox]::Show(
        $(New-Object -TypeName System.Windows.Forms.Form -Property @{ TopMost = $true }),
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null
}

function Find-File {
    param (
        [string]$Pattern,
        [string]$Pattern2 = "",
        [string]$Directory = "C:\Users\Hirossport\Desktop\",
        [bool]$PromptIfNotFound = $false,
        [string]$PromptTitle = "Fájl kiválasztása"
    )
    $items = "" -eq $Pattern2 ? (Get-ChildItem $Directory | Where-Object { $_.Name.Normalize() -match $Pattern } | Sort-Object -Property Name -Descending | Select-Object -First 1 ) : (Get-ChildItem $Directory | Where-Object { $_.Name.Normalize() -match $Pattern } | Where-Object Name -match $Pattern2 | Sort-Object -Property Name -Descending | Select-Object -First 1)
       
    if ($items) {
        return $items.FullName
    } elseif ($PromptIfNotFound) {
        Show-ErrorMessage -Message "Nem találom a következőt: $PromptTitle"

        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $Directory
        $OpenFileDialog.filter = "Minden Excel fájl|*.xls;*.xlsx;*.xlsm"
        $OpenFileDialog.Title = $PromptTitle
        $OpenFileDialog.ShowDialog() |  Out-Null
        $items = Get-ChildItem $Directory | Where-Object Name -match $OpenFileDialog.SafeFileName
        return $items.FullName
    }
    return ""
}

function get-Hours {
    param (
    [string]
    #String, mely az igényekből tartalmazza a műszak leírását
    $definition
    ) 
    $definition = $definition.Replace(' ', '')
    if ($definition -match '^\d{2}:\d{2}-\d{2}:\d{2}' -or
        $definition -match '^\d{1}:\d{2}-\d{2}:\d{2}' -or
        $definition -match '^\d{2}:\d{2}-\d{1}:\d{2}') {
        $start = $yesterday.Date
        $start = $start.AddHours($definition.Split("-")[0].Split(":")[0])
        $start = $start.AddMinutes($definition.Split("-")[0].Split(":")[1])

        $end = $yesterday.Date
        $end = $end.AddHours($definition.Split("-")[1].Split(":")[0])
        $end = $end.AddMinutes($definition.Split("-")[1].Split(":")[1])
        if ($end -lt $start) {
            $end = $end.AddDays(1)
        }
        $diff = $end.Subtract($start)
        $shift = [PSCustomObject]@{
            definition     = $start.ToString("HH:mm") + "-" +  $end.ToString("HH:mm")
            hours = ($diff.TotalMinutes % 60 -eq 0) ? $diff.Hours : $diff.TotalMinutes/60  
        }
        return $shift
    }
    elseif ($definition -eq "12óra") {
        get-Hours("08:30-20:30")
    }
    elseif ($definition -eq "2+4é") {
        get-Hours ('13:30-01:30')
    }
    elseif ([regex]::Matches($definition, '(?<=\().+?(?=\))' ).Success) {
        get-Hours(([regex]::Matches($definition, '(?<=\().+?(?=\))' ).Value))
    }
    else {
        throw "Rossz formátum"
    }
     <#
        .SYNOPSIS

        Értelmezi az igényben található munkaidő leírást

        .DESCRIPTION

        Párt képez, melíy tartalmazza, hogy mettől meddig tart egy műszak (String), és hogy az hány órát jelent (Int)
        Rekurzív függvény, mely első körben a leírás stringből kiszedi azt, hogy meddig tart egy műszak a következő formátumban ÓÓ:PP-ÓÓ:PP, majd második körben értelmezi azt.
    #>
}

function get-IgenyekHelper {
    param (
        $Map,
        [pscustomobject]
        $igeny
    )
    $total = 0    
    $limit = $igeny.Worksheet.Cells.Find("Mindösszesen").row
    for ($i = 4; $i -lt $limit; $i++)
    {
        $hours = $igeny.Worksheet.Cells($i, 2 + $yesterday.Day).Value2
        if ($null -ne $hours)
        {   
            $role =  [string] $igeny.Worksheet.Cells($i, 1).MergeArea().Value2
            try {
                $pair = get-Hours $igeny.Worksheet.Cells($i, 2).Value2
            }
            catch {
                $igeny.Workbook.close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($igeny.Application)
                exit
            }
            for ($j = 0; $j -lt $hours / $pair.hours; $j++) {
                $Map[$role]+=,$pair
                $total += $pair.hours
            }                 
        }
    }
    if ($igeny.Worksheet.Cells($limit, 2 + $yesterday.Day).Value2 -ne $total) {
        $Map
        Write-Output "Valami nem jó, nem egyezik az elszámolásba beírt órák száma az igénnyel"         
        $igeny.Application.visible=$true
        Show-ErrorMessage -Message "$total Valami nem jó, nem egyezik az elszámolásba beírt órák száma az igénnyel. `nKérlek nézd át!"
        exit
    }
    $igeny.Workbook.close($false)
}

function get-Igenyek {
    param (
        $muszakok,
        [string]
        $UszomesterIgenyPath,
        [string]
        $FrontOfficeIgenyPath
    )
    $excelHelper = New-Object -comobject Excel.Application
    $excelHelper.visible = $false
    get-IgenyekHelper  $muszakok @{
        "Application" = $excelHelper
        "Workbook"    = $excelHelper.Workbooks.Open($UszomesterIgenyPath)
        "Worksheet"   = $excelHelper.Workbooks.Open($UszomesterIgenyPath).Sheets.Item(1)
    } | Out-Null
    get-IgenyekHelper $muszakok  @{
        "Application" = $excelHelper
        "Workbook"    = $excelHelper.Workbooks.Open($FrontOfficeIgenyPath)
        "Worksheet"   = $excelHelper.Workbooks.Open($FrontOfficeIgenyPath).Sheets.Item(1)
    } | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelHelper)
}

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
    $PathFU = Find-File -Pattern $PatternFU -Pattern2 "^((?!Front).)*$" -PromptIfNotFound $FillCompletely -PromptTitle "Fürdő úszómester igényei"

    $patternfp =  '^F.*(' + [regex]::escape($honap) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month) +').*Fürdő.*(azd|ront).*xlsx$'
    $PathFP = Find-File -Pattern $patternfp -PromptIfNotFound $FillCompletely -PromptTitle "Fürdő front office igényei"

    if (-not $PathFU -or -not $PathFP) {
        $Empty = @{ definition = ""; hours = "" }
        $igeny = [ordered] @{"Úszómester" = @($empty;$empty)
        "Csúszda" = @($empty)
        "Gyógymedence`nfelügyelet" = @($empty;$empty)
        "Öltöző szolgálat" = @($empty;$empty)
        Gyermekmegőrző = @($empty)
        "Pénztáros" = @($empty;$empty)
        Recepció = @($empty)
    }
    } else {
        $Igeny = [ordered]@{}
        Get-Igenyek -Muszakok $Igeny -UszomesterIgenyPath $PathFU -FrontOfficeIgenyPath $PathFP
    }
    
    
    if (0 -ne $igeny.Count) {
        $furdosablon = Create-Template $igeny $fillCompletely $("Napi diákmunka elszámolás " + $yesterday.ToString("yyyy. MMMM dd."))
        if (-Not (Test-Path $savaAsFurdo)) {
            $furdosablon.Workbook.SaveAs($savaAsFurdo,[Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
        }
        $furdosablon.Application.visible=$true
        #Set-ItemProperty -Path HKCU:\Software\Script -Name YesterdaysWorkingHours -value $Today.Day
    }
}

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
    $PathSU = Find-File -Pattern $PatternSU  -Pattern2 "^((?!Front).)*$" -PromptIfNotFound $FillCompletely -PromptTitle "Strand úszómester igényei"
    
    $patternsp = '^F.*(' + [regex]::escape($honap) + '|' + [regex]::escape((Get-Culture).DateTimeFormat.GetMonthName($yesterday.Month)) +').*Strand.*(azd|ront).*xlsx$'
    $PathSP = Find-File -Pattern $PatternSP -PromptIfNotFound $FillCompletely -PromptTitle "Strand front office igényei"
    
    if (-not $PathSU -or -not $PathSP) {
        $Empty = @{ definition = ""; hours = "" }
        $Igeny = [ordered]@{
            Úszómester = @($Empty; $Empty; $Empty)
            Csúszda = @($Empty)
            "Vízőr (pancsoló)" = @($Empty)
            "Gyógymedence" = @($Empty)
            "Öltöző női" = @($Empty)
            "Öltöző férfi" = @($Empty)
            Kisegítő = @($Empty)
            Pénztáros = @($Empty)
        }
    } else {
        $Igeny = [ordered]@{}
        Get-Igenyek -Muszakok $Igeny -UszomesterIgenyPath $PathSU -FrontOfficeIgenyPath $PathSP
    }

    $saveAsStrand= "$($saveAsDir)$($yesterday.ToString('yyyy.MM.dd.')) - STRAND.xlsx"  
    
    if (0 -ne $igeny.Count) {
        $strandsablon = Create-Template $igeny $fillCompletely $("Napi diákmunka elszámolás - STRAND - " + $yesterday.ToString("yyyy. MMMM dd."))
        if (-Not (Test-Path $saveAsStrand)) {
            $strandsablon.Workbook.SaveAs($saveAsStrand,[Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, $true)
        }
        $strandsablon.Application.visible = $true
    }
    #Set-ItemProperty -Path HKCU:\Software\Script -Name YesterdaysWorkingHours -value $Today.Day
}