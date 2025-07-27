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

    $TimeRanges = Get-BlockedPoolTimeRanges -Sheet $workbook.Sheets.Item(1) -Pools $pools

    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    return $TimeRanges
}

function Set-Triggers {
    # Get the pool occupancy data inside the function
    $Pools = Get-Occupancies
    $TaskFolder = "\"  # Default to root folder in Task Scheduler

    foreach ($pool in $Pools) {
        $poolName = $pool.PoolName
        $blockedRanges = $pool.BlockedRanges

        if (-not $blockedRanges -or $blockedRanges.Count -eq 0) {
            Write-Verbose "Skipping pool $poolName because it has no blocked ranges."
            continue
        }

        $occTaskName = "${poolName}_Occ"
        $freeTaskName = "${poolName}_Free"

        $occTriggers = @()
        $freeTriggers = @()

        foreach ($range in $blockedRanges) {
            $start = [datetime]$range.Start
            $end = [datetime]$range.End

            $occTriggers += New-ScheduledTaskTrigger -Once -At $start
            $freeTriggers += New-ScheduledTaskTrigger -Once -At $end
        }

        function Update-TaskTriggers($taskName, $triggers) {
            Write-Verbose "Updating triggers for task '$taskName' in folder '$TaskFolder'"

            try {
                # You may want to verify the task exists before setting triggers
                $task = Get-ScheduledTask -TaskName $taskName -TaskPath $TaskFolder -ErrorAction Stop

                if ($null -eq $task) {
                    Write-Warning "Task '$taskName' was not found."
                    return
                }

                # Simplified trigger update command as you requested:
                Set-ScheduledTask -TaskName $taskName -TaskPath $TaskFolder -Trigger $triggers

                Write-Host "Updated triggers for task $taskName"
            }
            catch {
                Write-Warning "Error updating task '$taskName': $($_.Exception.Message)"
            }
        }

        Update-TaskTriggers -taskName $occTaskName -triggers $occTriggers
        Update-TaskTriggers -taskName $freeTaskName -triggers $freeTriggers
    }
}

