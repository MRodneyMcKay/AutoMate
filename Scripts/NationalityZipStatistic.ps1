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
$resolvedModulePath = (Resolve-Path -Path $ModulePath).Path

Import-Module (Join-Path -Path $resolvedModulePath -ChildPath 'VisitorStatistic\VisitorStatistic.psd1')

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

# Function to generate Excel reports
function New-ExcelReport {
    param (
        [string]$FilePath,
        [string]$SheetName,
        [array]$Headers,       # Column headers
        [array]$Data,          # Data array (PSCustomObjects)
        [hashtable]$Summary    # Summary data
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = $SheetName

    # Write headers
    for ($col = 0; $col -lt $Headers.Count; $col++) {
        $sheet.Cells.Item(1, $col + 1).Value2 = $Headers[$col]
    }

    # Write data rows dynamically based on headers
    $row = 2
    foreach ($item in $Data) {
        for ($col = 0; $col -lt $Headers.Count; $col++) {
            $propertyName = $Headers[$col] # Use the header name to access the property dynamically
            $sheet.Cells.Item($row, $col + 1).Value2 = [string]$item.$propertyName
            if ($propertyName -eq "Darabszám") {
                $sheet.Cells.Item($row, $col + 1).NumberFormat = '# ##0" fő"' # Format numeric columns
            }
        }
        $row++
    }
    $Summary
    # Write summary rows in order
    $summaryRow = 2
    $orderedKeys = $Summary.Keys 
    foreach ($key in $orderedKeys) {
        $sheet.Range("E$summaryRow").Value2 = $key
        $sheet.Range("F$summaryRow").Value2 = [string]$Summary[$key]
        $sheet.Range("F$summaryRow").NumberFormat = '# ##0" fő"' # Format the summary values as numbers
        $summaryRow++
    }

    # Auto-fit columns
    $sheet.Columns.AutoFit()

    # Save and release resources
    $workbook.SaveAs($FilePath)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}

# Main Script
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Válassz egy mappát, ahol a fájlok találhatók:"
$folderBrowser.ShowNewFolderButton = $false

if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedPath = $folderBrowser.SelectedPath
    $pairs = Get-MonthlyFilePairs -Path $selectedPath

    foreach ($pair in $pairs) {
        # Generate nationality report
        $nationalities = Get-GroupedNationalityData -csvPath $pair.NemzPath
        if (-not $nationalities -or $nationalities.Count -eq 0) {
            Write-Warning "No valid nationality data found for $($pair.NemzPath). Skipping report generation."
            continue
        }

        $nationalStats = Get-NationalityStats $nationalities
        $nationalSummary = [ordered]@{
            "Külföldi ország EU-n belül" = $nationalStats.Kulfold_EU
            "Külföldi ország EU-n kívül" = $nationalStats.Kulfold_NonEU
            "Összesen:"                  = $nationalStats.Kulfold_Osszesen
        }

        New-ExcelReport     -FilePath "$selectedPath\$($pair.Month) - Nemeztiségek.xlsx" `
                            -SheetName "$($pair.Month) - Nemeztiségek" `
                            -Headers @("ISO", "Ország", "Darabszám") `
                            -Data $nationalities `
                            -Summary $nationalSummary

        # Generate zip code report
        $zipCodes = Get-CleanedUpZipCodes -csvPath $pair.IrszPath -DesiredTotal (($nationalities | Where-Object { $_.ISO -eq "HU" }).Darabszám)
        if (-not $zipCodes -or $zipCodes.Count -eq 0) {
            Write-Log "No valid zip code data found for $($pair.IrszPath). Skipping report generation." WARNING
            continue
        }

        $zipStats = Get-ZipStats $zipCodes
        $zipSummary = [ordered]@{
            "Kecskemét"                             = $zipStats.Kecskemet
            "Kecskemét környéke"                    = $zipStats.Kecskemet_kornyeke
            "Egyéb magyarországon belüli település" = $zipStats.Egyeb_telepules
            "Összesen"                              = $zipStats.Osszesen
        }

        New-ExcelReport     -FilePath "$selectedPath\$($pair.Month) - Irányítószámok.xlsx" `
                            -SheetName "$($pair.Month) - Irányítószámok" `
                            -Headers @("Irányítószám", "Település", "Darabszám") `
                            -Data $zipCodes `
                            -Summary $zipSummary
    }
} else {
    Write-LOG "Nincs mappa kiválasztva." ERROR
}
