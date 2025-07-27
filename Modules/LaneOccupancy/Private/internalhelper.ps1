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

function Get-TodaysScheduleFile {
     param (
        $today = (Get-Date)
    )
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    

    # Get all matching files on Desktop
    $files = Get-ChildItem -Path $desktopPath -Filter "Pálya *.xlsx" | Where-Object {
        $_.BaseName -match "^Pálya .+"
    }

    foreach ($file in $files) {
        # Strip prefix and suffix
        $namePart = $file.BaseName -replace "^Pálya ", "" -replace "\.xlsx$", ""

        # Try full cross-year: YYYY.MM.dd.-YYYY.MM.dd.
        if ($namePart -match "^(\d{4})\.(\d{2})\.(\d{2})-(\d{4})\.(\d{2})\.(\d{2})\.$") {
            $startDate = Get-Date -Year $matches[1] -Month $matches[2] -Day $matches[3]
            $endDate   = Get-Date -Year $matches[4] -Month $matches[5] -Day $matches[6]
        }
        # Try same year, full month-day: YYYY.MM.dd.-MM.dd.
        elseif ($namePart -match "^(\d{4})\.(\d{2})\.(\d{2})-(\d{2})\.(\d{2})\.$") {
            $year = [int]$matches[1]
            $startDate = Get-Date -Year $year -Month $matches[2] -Day $matches[3]
            $endDate   = Get-Date -Year $year -Month $matches[4] -Day $matches[5]
        }
        # Try minimal form: YYYY.MM.dd.-DD.
        elseif ($namePart -match "^(\d{4})\.(\d{2})\.(\d{2})-(\d{2})\.$") {
            $year = [int]$matches[1]
            $month = [int]$matches[2]
            $startDate = Get-Date -Year $year -Month $month -Day $matches[3]
            $endDate   = Get-Date -Year $year -Month $month -Day $matches[4]
        }
        else {
            continue # Doesn't match any known pattern
        }

        if ($today -ge $startDate -and $today -le $endDate) {
            return $file.FullName
        }
    }

    throw "No matching schedule file found for today's date ($($today.ToShortDateString()))."
}

function Get-PoolRanges {
    param (
        [Parameter(Mandatory = $true)]
        $Sheet,

        [Parameter()]
        $Date = (Get-Date)
    )

    # Hungarian weekday name
    $dayMap = @{
        Monday    = "Hétfő"
        Tuesday   = "Kedd"
        Wednesday = "Szerda"
        Thursday  = "Csütörtök"
        Friday    = "Péntek"
        Saturday  = "Szombat"
        Sunday    = "Vasárnap"
    }
    $todayNameHu = $dayMap[$Date.DayOfWeek.ToString()]

    $rowCount = $Sheet.UsedRange.Rows.Count
    $timeCols = @(1, 13, 21)         # A, M, U
    $lastLaneCols = @(11, 19, 23)    # K, S, W
    $startRow = $null
    $endRow = $null

    for ($row = 1; $row -le $rowCount; $row++) {
        for ($i = 0; $i -lt $timeCols.Count; $i++) {
            $timeVal = $Sheet.Cells.Item($row, $timeCols[$i]).Text
            $poolName = $Sheet.Cells.Item($row, $lastLaneCols[$i]).Text
            if ($timeVal -like "$todayNameHu*" -and $poolName -ne "") {
                $startRow = $row + 2
                break
            }
        }
        if ($startRow) { break }
    }

    if (-not $startRow) {
        Write-Error "Nem található a(z) $todayNameHu napi blokk."
        return
    }

    for ($row = $startRow; $row -le $rowCount; $row++) {
        foreach ($col in $timeCols) {
            $val = $Sheet.Cells.Item($row, $col).Text
            if ($val -match "-\s*21:00$") {
                $endRow = $row
                break
            }
        }
        if ($endRow) { break }
    }

    $rangeInfo = [PSCustomObject]@{
        StartRow = $startRow
        EndRow   = $endRow
    }
    return @(
    [PSCustomObject]@{
        Name     = "Versenymedence, 50 m"
        StartRow = $rangeInfo.StartRow
        EndRow   = $rangeInfo.EndRow
        StartCol = 1    # A
        EndCol   = 11   # K
    },
    [PSCustomObject]@{
        Name     = "Bemelegítőmedence, 25m"
        StartRow = $rangeInfo.StartRow
        EndRow   = $rangeInfo.EndRow
        StartCol = 13   # M
        EndCol   = 19   # S
    },
    [PSCustomObject]@{
        Name     = "Tanmedence"
        StartRow = $rangeInfo.StartRow
        EndRow   = $rangeInfo.EndRow
        StartCol = 21   # U
        EndCol   = 23   # W
    },
    [PSCustomObject]@{
        Name     = "Kalandmedence"
        StartRow = $rangeInfo.StartRow
        EndRow   = $rangeInfo.EndRow
        StartCol = 25   # Y
        EndCol   = 27   # AA
    }
)
}