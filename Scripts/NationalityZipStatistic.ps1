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

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

Import-Module (Join-Path -Path $resolvedModulePath -ChildPath 'VisitorStatistic\VisitorStatistic.psd1')
Import-Module (Join-Path -Path $resolvedModulePath -ChildPath 'WPFToolkit\WPFToolkit.psd1')

[System.Reflection.Assembly]::LoadFrom([System.Environment]::GetEnvironmentVariable("OfficeAssemblies_Excel", [System.EnvironmentVariableTarget]::User))
Add-Type -AssemblyName System.Windows.Forms

. "$PSScriptRoot\ViewModels\NationalityZipStatisticViewModel.ps1"
. "$PSScriptRoot\Views\NationalityZipStatisticView.ps1"

function Get-MonthlyFilePairs {
    param (
        [string[]]$NationalityFiles,
        [string[]]$ZipFiles
    )

    $filesByMonth = @{
        Nationality = @{}
        Zip         = @{}
    }

    foreach ($fileGroup in @(
            @{ Type = 'Nationality'; Paths = $NationalityFiles },
            @{ Type = 'Zip'; Paths = $ZipFiles }
        )) {
        foreach ($path in $fileGroup.Paths) {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "A fájl nem található: $path"
            }
            if ([System.IO.Path]::GetExtension($path) -ine '.csv') {
                throw "Csak CSV-fájl használható: $path"
            }

            $prefix = Get-FilePrefix $path $fileGroup.Type
            if ($filesByMonth[$fileGroup.Type].ContainsKey($prefix)) {
                throw "Több $($fileGroup.Type) fájl van ehhez az előtaghoz: $prefix"
            }
            $filesByMonth[$fileGroup.Type][$prefix] = $path
        }
    }

    $prefixes = @(
        $filesByMonth.Nationality.Keys
        $filesByMonth.Zip.Keys
    ) | Sort-Object -Unique
    if ($prefixes.Count -eq 0) {
        throw "Válasszon legalább egy nemzetiségi és egy irányítószám-fájlt."
    }

    $pairs = foreach ($prefix in $prefixes) {
        if (-not $filesByMonth.Nationality.ContainsKey($prefix) -or
            -not $filesByMonth.Zip.ContainsKey($prefix)) {
            throw "Hiányzik az előtaghoz tartozó párfájl: $prefix"
        }

        [PSCustomObject]@{
            Prefix   = $prefix
            IrszPath = $filesByMonth.Zip[$prefix]
            NemzPath = $filesByMonth.Nationality[$prefix]
        }
    }

    return @($pairs)
}

function Get-FilePrefix {
    param (
        [string]$Path,
        [string]$Type
    )

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $marker = if ($Type -eq "Nationality") { "nemzet" } else { "irányító" }
    $markerIndex = $fileName.IndexOf($marker, [System.StringComparison]::OrdinalIgnoreCase)
    if ($markerIndex -lt 0) {
        throw "A fájlnév nem tartalmazza a várt kifejezést '$marker': $fileName"
    }

    $prefix = $fileName.Substring(0, $markerIndex).TrimEnd([char[]]@(' ', '-', '_', '.'))
    if ([string]::IsNullOrWhiteSpace($prefix)) {
        throw "Nem olvasható ki előtag a fájlnévből: $fileName"
    }

    return $prefix
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
    $sheet.Columns.AutoFit() | Out-Null

    # Save and release resources
    $workbook.SaveAs($FilePath) | Out-Null
    $workbook.Close($false) | Out-Null
    $excel.Quit() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}





function Invoke-NationalityZipStatistics {
    param (
        [string[]]$NationalityFiles,
        [string[]]$ZipFiles,
        [string]$OutputDirectory,
        [scriptblock]$ProgressCallback
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        throw "A mentési mappa nem található: $OutputDirectory"
    }

    $pairs = Get-MonthlyFilePairs -NationalityFiles $NationalityFiles -ZipFiles $ZipFiles
    $completedPairs = 0
    $pairPrefixes = [System.Collections.Generic.List[string]]::new()
    for ($pairIndex = 0; $pairIndex -lt $pairs.Count; $pairIndex++) {
        $pair = $pairs[$pairIndex]
        if ($null -ne $ProgressCallback) {
            & $ProgressCallback $pair.Prefix ($pairIndex + 1) $pairs.Count
        }

        $nationalities = Get-GroupedNationalityData -csvPath $pair.NemzPath
        if (-not $nationalities -or $nationalities.Count -eq 0) {
            throw "Nem található érvényes nemzetiségi adat: $($pair.NemzPath)"
        }

        $nationalStats = Get-NationalityStats $nationalities
        $nationalSummary = [ordered]@{
            "Külföldi ország EU-n belül" = $nationalStats.Kulfold_EU
            "Külföldi ország EU-n kívül" = $nationalStats.Kulfold_NonEU
            "Összesen:"                  = $nationalStats.Kulfold_Osszesen
        }

        $nationalityReportPath = Join-Path $OutputDirectory "$($pair.Prefix) - Nemzetiségek.xlsx"
        New-ExcelReport -FilePath $nationalityReportPath `
            -SheetName "$($pair.Prefix) - Nemzetiségek" `
            -Headers @("ISO", "Ország", "Darabszám") `
            -Data $nationalities `
            -Summary $nationalSummary | Out-Null

        $desiredTotal = ($nationalities | Where-Object { $_.ISO -eq "HU" }).Darabszám
        $zipCodes = Get-CleanedUpZipCodes -csvPath $pair.IrszPath -DesiredTotal $desiredTotal
        if (-not $zipCodes -or $zipCodes.Count -eq 0) {
            throw "Nem található érvényes irányítószám-adat: $($pair.IrszPath)"
        }

        $zipStats = Get-ZipStats $zipCodes
        $zipSummary = [ordered]@{
            "Kecskemét"                             = $zipStats.Kecskemet
            "Kecskemét környéke"                    = $zipStats.Kecskemet_kornyeke
            "Egyéb magyarországon belüli település" = $zipStats.Egyeb_telepules
            "Összesen:"                             = $zipStats.Osszesen
        }

        $zipReportPath = Join-Path $OutputDirectory "$($pair.Prefix) - Irányítószámok.xlsx"
        New-ExcelReport -FilePath $zipReportPath `
            -SheetName "$($pair.Prefix) - Irányítószámok" `
            -Headers @("Irányítószám", "Település", "Darabszám") `
            -Data $zipCodes `
            -Summary $zipSummary | Out-Null

        $completedPairs++
        $pairPrefixes.Add($pair.Prefix)
    }

    return [PSCustomObject]@{
        PairCount = $completedPairs
        Prefixes  = [string[]]$pairPrefixes
    }
}

$app = [System.Windows.Application]::Current
if (-not $app) {
    $app = [System.Windows.Application]::new()
    $app.ShutdownMode = 'OnExplicitShutdown'
    $app.Resources.MergedDictionaries.Add((Get-ThemeResources))
    $app.Resources.MergedDictionaries.Add((Get-ThemeStyle))
}

$viewModel = [NationalityZipStatisticViewModel]::new({
        param($nationalityFiles, $zipFiles, $outputDirectory, $progressCallback)
        Invoke-NationalityZipStatistics `
            -NationalityFiles $nationalityFiles `
            -ZipFiles $zipFiles `
            -OutputDirectory $outputDirectory `
            -ProgressCallback $progressCallback
    })
$view = [NationalityZipStatisticView]::new($viewModel)
$frame = [System.Windows.Threading.DispatcherFrame]::new()
$view.Window.Add_Closed({
        $frame.Continue = $false
    }.GetNewClosure())
$view.Show()
[System.Windows.Threading.Dispatcher]::PushFrame($frame)

Remove-Variable app, viewModel, view, frame -ErrorAction SilentlyContinue
[GC]::Collect()
