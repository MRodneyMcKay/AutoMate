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

$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$resolvedModulegPath = (Resolve-Path -Path $ModulePath).Path

Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'VisitorStatistic\VisitorStatistic.psd1')

Add-Type -AssemblyName System.Windows.Forms

function Get-MonthlyFilePairs {
    param (
        [string]$Path
    )

    $files = Get-ChildItem -Path $Path -Filter "*.csv" -File

    $parsed = foreach ($file in $files) {
        if ($file.Name -match '^(\d{4})\D?(\d{2}).*?(irányító|irányitoszám|nemzet).*?\.csv$') {
            [PSCustomObject]@{
                Year  = $matches[1]
                Month = $matches[2]
                Type  = $matches[3]
                Path  = $file.FullName
                Key   = "$($matches[1])-$($matches[2])"
            }
        }
    }
    $grouped = $parsed | Group-Object Key | Where-Object {
        $_.Group.Type -contains 'irányító' -or $_.Group.Type -contains 'irányitoszám'
    } | Where-Object {
        $_.Group.Type -contains 'nemzet'
    }
    $pairs = foreach ($group in $grouped) {
        $irsz = $group.Group | Where-Object { $_.Type -like 'irányító' -or $_.Type -like 'irányitoszám' }
        $nemz = $group.Group | Where-Object { $_.Type -eq 'nemzet' }

        [PSCustomObject]@{
            Month     = $group.Name
            IrszPath  = [string]$irsz.Path
            NemzPath  = [string]$nemz.Path
        }
    }

    return $pairs
}

# Example usage:

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Válassz egy mappát, ahol a fájlok találhatók:"
$folderBrowser.ShowNewFolderButton = $false
if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedPath = $folderBrowser.SelectedPath
    $pairs = Get-MonthlyFilePairs -Path $selectedPath
} else {
    Write-Host "Nincs mappa kiválasztva." -ForegroundColor Yellow
}

foreach ($pair in $pairs) {
    $pair.Month
    $nationalities = Get-GroupedNationalityData -csvPath  $pair.NemzPath
    $nationalStats = Get-NationalityStats $nationalities
    $zipCodes = Get-CleanedUpZipCodes -csvPath $pair.IrszPath -DesiredTotal (($nationalities | Where-Object { $_.ISO -eq "HU" }).Darabszám)
    $zipStats = Get-ZipStats $zipCodes
     # Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false 

# Add a new workbook
$workbook = $excel.Workbooks.Add()

# Get the first worksheet and rename it
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = "$($pair.Month) - Nemeztiségek"

# Write headers
$headers = @("ISO", "Ország", "Darabszám", "", "Megnevezés", "Darabszám")
for ($col = 0; $col -lt $headers.Count; $col++) {
    $sheet.Cells.Item(1, $col + 1).Value2 = $headers[$col]
}

# Write data rows
$row = 2
foreach ($item in $nationalities) {
    $sheet.Cells.Item($row, 1).Value2 = $item.ISO
    $sheet.Cells.Item($row, 2).Value2 = $item.Orszag
    $sheet.Cells.Item($row, 3).Value2 = [string]$item.Darabszám
    $sheet.Cells.Item($row, 3).NumberFormat = '# ##0" fő"'
    $row++
}

# Fill in descriptive labels
$sheet.Range("E2").Value2 = "Külföldi ország EU-n belül"
$sheet.Range("E3").Value2 = "Külföldi ország EU-n kívül"
$sheet.Range("E4").Value2 = "Összesen:"

# Fill in values from nationalStats
$sheet.Range("F2").Value2 = [string]$nationalStats.Kulfold_EU
$sheet.Range("F2").NumberFormat = '# ##0" fő"'
$sheet.Range("F3").Value2 = [string]$nationalStats.Kulfold_NonEU
$sheet.Range("F3").NumberFormat = '# ##0" fő"'
$sheet.Range("F4").Value2 = [string]$nationalStats.Kulfold_Osszesen
$sheet.Range("F4").NumberFormat = '# ##0" fő"'

# Optional: Auto-fit columns
$sheet.Columns.AutoFit()

$filePath = "$selectedPath\$($pair.Month) - Nemeztiségek.xlsx"
$workbook.SaveAs($filePath)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

    # Start Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false 
    
    # Add a new workbook
    $workbook = $excel.Workbooks.Add()
    
    # Get the first worksheet and rename it
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "$($pair.Month) - Irányítószámok"
# Write headers
$headers = @("Irányítószám", "Település", "Darabszám", "", "Megnevezés", "Darabszám")
for ($col = 0; $col -lt $headers.Count; $col++) {
    $sheet.Cells.Item(1, $col + 1).Value2 = $headers[$col]
}

# Write data rows
$row = 2
foreach ($item in $zipCodes) {
    $sheet.Cells.Item($row, 1).Value2 = $item.Zip
    $sheet.Cells.Item($row, 2).Value2 = $item.Megnevezes
    $sheet.Cells.Item($row, 3).Value2 = [string] $item.Darabszám
    $sheet.Cells.Item($row, 3).NumberFormat = '# ##0" fő"'
    $row++
}

# Fill in descriptive labels
$sheet.Range("E2").Value2 = "Kecskemét"
$sheet.Range("E3").Value2 = "Kecskemét környéke"
$sheet.Range("E4").Value2 = "Egyéb magyarországon belüli település"
$sheet.Range("E5").Value2 = "Összesen"

# Fill in values from nationalStats
$sheet.Range("F2").Value2 = [string]$zipStats.Kecskemet
$sheet.Range("F2").NumberFormat = '# ##0" fő"'
$sheet.Range("F3").Value2 = [string]$zipStats.Kecskemet_kornyeke
$sheet.Range("F3").NumberFormat = '# ##0" fő"'
$sheet.Range("F4").Value2 = [string]$zipStats.Egyeb_telepules
$sheet.Range("F4").NumberFormat = '# ##0" fő"'
$sheet.Range("F5").Value2 = [string]$zipStats.Osszesen
$sheet.Range("F5").NumberFormat = '# ##0" fő"'

    $sheet.Columns.AutoFit()

 $filePath = "$selectedPath\$($pair.Month) - Irányítószámok.xlsx"
$workbook.SaveAs($filePath)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) 

}
