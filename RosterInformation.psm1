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

# Load the Microsoft.Office.Interop.Excel assembly

$possiblePaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\Microsoft.Office.Interop.Excel.dll",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\Microsoft.Office.Interop.Excel.dll",
    "C:\Program Files\Microsoft Office\Office16\Microsoft.Office.Interop.Excel.dll",
    "C:\Program Files (x86)\Microsoft Office\Office16\Microsoft.Office.Interop.Excel.dll",
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\*\Microsoft.Office.Interop.Excel.dll"
)

$dllPath = $possiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($dllPath) {
    Add-Type -Path $dllPath
} else {
    Write-Warning "Could not find Microsoft.Office.Interop.Excel in PS7!"
}

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
        Write-Error "No matching files found."
        return
    }

    $filePath = $items.FullName
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($filePath)
    $worksheet = $workbook.Sheets.Item(1)

    # Initialize an empty array
    $roster = @()

    try {        
        $roster += Get-Shift $worksheet 'Konfár Nikolett' $today
        $roster += Get-Shift $worksheet 'Raduska Zsolt' $today
    } catch {
        Write-Error "An error occurred: $_"
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

function Get-Shift {
    param (
        [object]$worksheet,
        [string]$name,
        [datetime]$date
    )
    $found = $worksheet.Cells.Find($name)
    if ($found) {
        return [PSCustomObject]@{
            Name  = $name
            Shift = $worksheet.Cells($found.Row, 3 + $date.Day).Text
        }
    } else {
        Write-Error "$name not found in the worksheet."
        return $null
    }
}

function Get-ShiftManager
{
    param (
        [datetime]$now = $(Get-Date)
    )
    $today = Get-Date $now
    $monthName = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName($today.Month))
    $pattern = '^(' + [regex]::escape($today.Year) +').*(' + [regex]::escape($monthName) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($today.Month) +')(?!.*(azd|ront|beosztás|havi)).*xlsx$'
    $items = Get-ChildItem $([System.Environment]::GetFolderPath("Desktop")) | Where-Object { $_.Name.Normalize() -match ($pattern) } | Sort-Object Name -Descending | Select-Object -First 1
    $filePath=$items.FullName
    $ExcelBack = New-Object -ComObject Excel.Application
    $ExcelBack.visible=$false
    Write-Log -Message "Opening file: $filePath"
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

function Get-Name {
    param (
        [object]$worksheet,
        [string]$ShiftID,
        [datetime]$date
    )
    $header = $workSheet.Cells.Find("műszakvezető")
    $ColumnOffset = $workSheet.Rows($header.Cells.Row() - 1).Find("sz").Cells.Column()
    $found = $workSheet.Range($workSheet.Cells($header.Cells.Row(), $($ColumnOffset + $date.Day)).Address(), $workSheet.Cells($( -1 + $header.Cells.Row() + $header.MergeArea.Rows.Count()), $($ColumnOffset + $date.Day)).Address()).Find($ShiftID, [Type]::Missing, [int][Microsoft.Office.Interop.Excel.XlFindLookIn]::xlValues, [int][Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole)
    if ($found) {
        return [PSCustomObject]@{
            Name  = $($workSheet.Cells($found.Row(), $header.Cells.Column + 1).Value2)
            Shift = $($workSheet.Cells($found.Row(), $($ColumnOffset + $date.Day)).Text)
        }
    } else {
        Write-Error "$ShiftID not found in the worksheet."
        return [PSCustomObject]@{
            Name  = $null
            Shift = $ShiftID
        }
    }
}