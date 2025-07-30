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

function Print-Today {

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $filename = Get-TodaysScheduleFile
    if (-not $filename) {
        "No schedule file found for today."
        return
    }

    $workbook = $excel.Workbooks.Open((Get-TodaysScheduleFile))
    $worksheet =  $workbook.Sheets.Item(1)
    $poolRanges = Get-PoolRanges -Sheet $worksheet | Select-Object -First 1
    $rangeInfo = "A$($poolRanges.StartRow-2):W$($poolRanges.EndRow)"
    $range =  $worksheet.Range($rangeInfo)
    $range.Select()
    $pageSetup = $worksheet.PageSetup
        if (-not $pageSetup) {
            throw "PageSetup object not accessible."
        }

        # Set print area and settings
        $pageSetup.PrintArea = $range.Address($false, $false)
        $pageSetup.FitToPagesWide = 1
        $pageSetup.FitToPagesTall = 1
        $pageSetup.Zoom = $false

        # Set margins to zero
        $pageSetup.LeftMargin   = 0
        $pageSetup.RightMargin  = 0
        $pageSetup.TopMargin    = 0
        $pageSetup.BottomMargin = 0

    # Print the range
    $worksheet.PrintOut()
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

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

