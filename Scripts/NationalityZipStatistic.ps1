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
            IrszPath  = $irsz.Path
            NemzPath  = $nemz.Path
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
    $nationalities = Get-GroupedNationalityData -Path $pair.NemzPath
    Get-NationalityStats $nationalities   
    $zipCodes = Get-CleanedUpZipCodes -Path $pair.IrszPath  ($nationalities | Where-Object { $_.ISO -eq "HU" }).Darabszám
    Get-ZipStats $zipCodes
}
