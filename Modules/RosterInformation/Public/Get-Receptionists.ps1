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

function Get-Receptionists {
    param (
        [datetime]$now = $(Get-Date)
    )

    $today = Get-Date $now
    $year = $today.Year
    $monthName = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName($today.Month))
    $pattern = '^(' + [regex]::escape($year) + ').*(' + [regex]::escape($monthName) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($today.Month) + ').*(azd|ront).*xlsx$'
    $items = Get-ChildItem $([System.Environment]::GetFolderPath("Desktop")) | Where-Object { $_.Name.Normalize() -match ($pattern) } | Sort-Object Name -Descending | Select-Object -First 1
    if ($items.Count -eq 0) {
        throw "Did not find front office schedule"
    }
    $filePath = $items.FullName
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($filePath)
    Write-Log -Message "Opened file: $filePath"
    $worksheet = $workbook.Sheets.Item(1)

    # Initialize an empty array
    $roster = @()

    try {        
        $roster += Get-Shift $worksheet 'Konfár Nikolett' $today
        $roster += Get-Shift $worksheet 'Raduska Zsolt' $today
    } catch {
        Write-Log $_.Exception.Message -Level "DEBUG"
        throw
    } finally {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$workbook)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$worksheet)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable excel -ErrorAction SilentlyContinue
    }

    return $roster
}