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
        Write-Log -Message "Fájl megtalálva: $($items.FullName)"
        return $items.FullName
    } elseif ($PromptIfNotFound) {
        Write-Log -Message "Nem találom a következőt: $PromptTitle" -Level "ERROR"

        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $Directory
        $OpenFileDialog.filter = "Minden Excel fájl|*.xls;*.xlsx;*.xlsm"
        $OpenFileDialog.Title = $PromptTitle
        $OpenFileDialog.ShowDialog() |  Out-Null
        $items = Get-ChildItem $Directory | Where-Object { $_.Name.Normalize() -eq $OpenFileDialog.SafeFileName.Normalize() }
        Write-Log -Message "Fájl kiválasztva: $($items.FullName)"
        return $items.FullName
    }
    Write-Log -Message "Nem találom a következőt: $PromptTitle" -Level "WARNING"
    throw [System.IO.IOException]::new("Nem találom a következőt: $PromptTitle") 
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
    elseif ($definition -eq "4é") {
        get-Hours("21:00-01:00")
    }
    elseif ($definition -eq "5é") {
        get-Hours("20:30-01:30")
    }
    elseif ($definition -eq "2+4é") {
        get-Hours ('13:00-00:30')
    }
    elseif ([regex]::Matches($definition, '(?<=\().+?(?=\))' ).Success) {
        get-Hours(([regex]::Matches($definition, '(?<=\().+?(?=\))' ).Value))
    }
    else {
        throw [System.ArgumentException]::new("Rossz formátum a műszak leírásának: $definition")
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
                throw
            }
            for ($j = 0; $j -lt $hours / $pair.hours; $j++) {
                $Map[$role]+=,$pair
                $total += $pair.hours
            }                 
        }
    }
    if ($igeny.Worksheet.Cells($limit, 2 + $yesterday.Day).Value2 -ne $total) {        
        $igeny.Application.visible=$true
        Write-Log -Message "$total Valami nem jó, nem egyezik az elszámolásba beírt órák száma az igénnyel. `nKérlek nézd át!" -Level "ERROR"           
        exit
    }
    else {
        Write-Log -Message "Az igényeket sikeresen beolvastam!"
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