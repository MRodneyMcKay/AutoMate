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

[System.Reflection.Assembly]::LoadFrom([System.Environment]::GetEnvironmentVariable("OfficeAssemblies_Excel", [System.EnvironmentVariableTarget]::User))
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
        [System.Collections.IDictionary]$Summary    # Summary data (preserves ordered hashtable)
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

            # Safe property retrieval
            $value = ""
            if ($null -ne $item -and $item.PSObject -and $item.PSObject.Properties.Match($propertyName)) {
                $value = $item.PSObject.Properties[$propertyName].Value
            }

            if ($propertyName -eq "Darabszám") {
                try {
                    if ($value -eq $null -or $value -eq "") {
                        $numeric = 0.0
                    } else {
                        $numeric = [double]$value
                    }
                } catch {
                    $numeric = 0.0
                }
                $sheet.Cells.Item($row, $col + 1).Value2 = [string]$numeric
                $sheet.Cells.Item($row, $col + 1).NumberFormat = '# ##0" fő"'
            } else {
                $sheet.Cells.Item($row, $col + 1).Value2 = [string]$value
            }
        }
        $row++
    }

    # Write summary rows in order (preserve ordered hashtable)
    $summaryRow = 2
    $startSummaryRow = 2 # Track the start of the summary region
    foreach ($entry in $Summary.GetEnumerator()) {
        $key = $entry.Key
        $val = $entry.Value

        $sheet.Range("E$summaryRow").Value2 = $key

        if ($val -is [int] -or $val -is [long] -or $val -is [double] -or $val -is [decimal]) {
            $sheet.Range("F$summaryRow").Value2 = [double]$val
        } else {
            if ($null -eq $val -or $val -eq "") {
                $sheet.Range("F$summaryRow").Value2 = ""
            } else {
                try {
                    $num = [double]::Parse([string]$val)
                    $sheet.Range("F$summaryRow").Value2 = $num
                } catch {
                    $sheet.Range("F$summaryRow").Value2 = [string]$val
                }
            }
        }

        $sheet.Range("F$summaryRow").NumberFormat = '# ##0" fő"'

        # Apply alternating row colors (correct range syntax)
        if (($summaryRow % 2) -eq 0) {
            $sheet.Range("E$summaryRow : F$summaryRow").Interior.Color = 0xE6E6E6
        } else {
            $sheet.Range("E$summaryRow : F$summaryRow").Interior.Color = 0xFFFFFF
        }

        # Apply a stronger color and double upper border for "Összesen"
        if ($key -eq "Összesen" -or $key -eq "Összesen:") {
            $sheet.Range("E$summaryRow : F$summaryRow").Interior.Color = 0xCCCCCC
            $sheet.Range("E$summaryRow : F$summaryRow").Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlDouble
        }

        $summaryRow++
    }

    $summaryRange = $sheet.Range("E2 : F$($summaryRow - 1)")
    $summaryRange.Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $summaryRange.Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeLeft).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium

    $summaryRange.Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $summaryRange.Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeRight).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium

    $summaryRange.Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $summaryRange.Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeTop).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium

    $summaryRange.Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
    $summaryRange.Borders.Item([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).Weight = [Microsoft.Office.Interop.Excel.XlBorderWeight]::xlMedium

    # Auto-fit columns
    $sheet.Columns.AutoFit()

    # Save and release resources
    $workbook.SaveAs($FilePath)
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
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
            "Összesen:"                             = $zipStats.Osszesen
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
