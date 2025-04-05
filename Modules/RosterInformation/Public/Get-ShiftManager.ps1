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

function Get-ShiftManager
{
    param (
        [datetime]$now = $(Get-Date)
    )
    $today = Get-Date $now
    $monthName = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName($today.Month))
    $pattern = '^(' + [regex]::escape($today.Year) +').*(' + [regex]::escape($monthName) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($today.Month) +')(?!.*(azd|ront|beosztás|havi)).*xlsx$'
    $items = Get-ChildItem $([System.Environment]::GetFolderPath("Desktop")) | Where-Object { $_.Name.Normalize() -match ($pattern) } | Sort-Object Name -Descending | Select-Object -First 1
    if ($items.Count -eq 0) {
        Write-Log -Message "No matching files found for staff schedule." -Level "ERROR"
        return
    }
    $filePath=$items.FullName
    $ExcelBack = New-Object -ComObject Excel.Application
    $ExcelBack.visible=$false
    $Workbook = $ExcelBack.Workbooks.Open($filePath)
    $workSheet = $Workbook.Sheets.Item(1)
    $roster = @()
    $roster += Get-Name $worksheet '1~*' $today
    $roster += Get-Name $worksheet '*2~*' $today
    $Workbook.close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ExcelBack)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable ExcelBack -ErrorAction SilentlyContinue
    return $roster
}