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

function Get-Shift {
    
    param (
        [object]$worksheet,
        [string]$name,
        [datetime]$date
    )
    $found = $worksheet.Cells.Find($name)
    if ($null -eq $found) {
        throw "$name not found in the worksheet."
        return $null
    }
    return [PSCustomObject]@{
        Name  = $name
        Shift = $worksheet.Cells($found.Row, 3 + $date.Day).Text
    }
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
        return $($workSheet.Cells($found.Row(), $header.Cells.Column + 1).Value2)
    } else {
        throw "$ShiftID not found in the worksheet."
    }
}