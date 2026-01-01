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

function Get-ShiftManager {
    param (
        [datetime]$now = (Get-Date)
    )

    # Use one consistent culture
    $culture = Get-Culture

    $today = Get-Date $now
    $year  = $today.Year
    $month = $culture.DateTimeFormat.GetMonthName($today.Month)

    # Escape everything used in regex
    $escapedYear  = [regex]::escape($year)
    $escapedMonth = [regex]::escape($month)

    # Build regex pattern
    $pattern = '^' +
        $escapedYear +
        '.*' +
        $escapedMonth +
        '(?!.*(azd|ront|beosztás|havi|Gyógyászat)).*\.xlsx$'

    # Normalize pattern once
    $normalizedPattern = $pattern.Normalize([Text.NormalizationForm]::FormC)

    # Find matching file
    $items = Get-ChildItem ([Environment]::GetFolderPath("Desktop")) |
        Where-Object {
            $_.Name.Normalize([Text.NormalizationForm]::FormC) -match $normalizedPattern
        } |
        Sort-Object Name -Descending |
        Select-Object -First 1

    if (-not $items) {
        throw "Did not find staff schedule"
    }

    $filePath = $items.FullName

    Write-Log "Using roster file: $filePath" -Level "DEBUG"

    # Open Excel
    $ExcelBack = New-Object -ComObject Excel.Application
    $ExcelBack.Visible = $false
    $Workbook = $ExcelBack.Workbooks.Open($filePath)
    $Worksheet = $Workbook.Sheets.Item(1)

    try {
        $name = Get-Name `
            $Worksheet `
            ($now.TimeOfDay -lt (New-TimeSpan -Hours 12 -Minutes 55) ? '1~*' : '*2~**') `
            $today
    }
    catch {
        Write-Log $_.Exception.Message -Level "DEBUG"
        throw
    }
    finally {
        $Workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ExcelBack)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable ExcelBack -ErrorAction SilentlyContinue
    }

    return $name
}
