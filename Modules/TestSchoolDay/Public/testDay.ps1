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

function Test-Schoolday {
    param (
        [datetime]$date = (Get-Date)
    )
    Write-Log -Message "Testing if $($date.ToString('yyyy-MM-dd')) is a school day."
    # Check if today is Saturday or Sunday
    if ($date.DayOfWeek -in @('Saturday', 'Sunday')) {
        Write-Log -Message "$($date.ToString('yyyy-MM-dd')) is a weekend."
        return $false
    }

    $schooldayRanges = @(
        [PSCustomObject]@{ Start = '2024-10-26'; End = '2024-11-03' },
        [PSCustomObject]@{ Start = '2024-12-21'; End = '2025-01-05' },
        [PSCustomObject]@{ Start = '2025-04-17'; End = '2025-04-27' },
        [PSCustomObject]@{ Start = '2025-06-21'; End = '2025-09-01' }
    )

    foreach ($range in $schooldayRanges) {
        if ($date -ge (Get-Date $range.Start) -and $date -le (Get-Date $range.End)) {
            Write-Log -Message "$($date.ToString('yyyy-MM-dd')) is within a schoolbreak"
            return $false
        }
    }

    Write-Log -Message "$($date.ToString('yyyy-MM-dd')) is a school day."
    return $true
}
