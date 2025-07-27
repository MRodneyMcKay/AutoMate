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

function Get-Occupancies {

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$filename = Get-TodaysScheduleFile
if (-not $filename) {
    "No schedule file found for today."
    return
}

$workbook = $excel.Workbooks.Open((Get-TodaysScheduleFile))

$pools = Get-PoolRanges -Sheet $workbook.Sheets.Item(1)

Get-BlockedPoolTimeRanges -Sheet $workbook.Sheets.Item(1) -Pools $pools

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
