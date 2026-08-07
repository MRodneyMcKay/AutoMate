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

class PrintTaskService {
    hidden [hashtable] $_Actions
    PrintTaskService() {
        $actionsPath = Join-Path $PSScriptRoot "PrintActions.ps1"
        if (-not (Test-Path $actionsPath)) {
            throw "PrintActions.ps1 not found: $actionsPath"
        }
        $PrintActions = $null
        . $actionsPath
        if ($null -eq $PrintActions) {
            throw "PrintActions.ps1 did not define `$PrintActions"
        }
        $this._Actions = $PrintActions
    }

    [void] Execute(
        [PrintTask] $task,
        [string] $sheetPath
    ) {
        if ($null -eq $task) {
            throw "Print task cannot be null"
        }
        if (-not $this._Actions.ContainsKey($task.ActionId)) {
            throw "No print action registered for ActionId '$($task.ActionId)'"
        }
        $action = $this._Actions[$task.ActionId]
        if ($task.RequiresSheetPath) {
            if ([string]::IsNullOrWhiteSpace($sheetPath)) {
                throw "Task '$($task.Name)' requires a sheet path"
            }
            Write-Host "Sheetpath: $sheetPath"
            & $action $sheetPath
        } else {
            & $action
        }
    }
}