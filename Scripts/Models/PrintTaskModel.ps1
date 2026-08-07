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

class PrintTask : WPFToolkit.MVVM.ObservableObject {

    [string] $Name
    [bool] $RequiresSheetPath
    [string] $ToolTip
    [string] $ActionId
    [WPFToolkit.MVVM.ObservableProperty[Nullable[bool]]] $IsSelected
    PrintTask(
        [string]$name,
        [bool]$requiresSheetPath,
        [string]$actionId
    ) {

        $this.Name = $name
        $this.RequiresSheetPath = $requiresSheetPath
        if ($this.RequiresSheetPath) {
            $this.ToolTip = "Ez a feladat egy kiválasztott jelenléti ív fájlt igényel."
        }
        else {
            $this.ToolTip = $null
        }
        $this.ActionId = $actionId
        $this.IsSelected = [WPFToolkit.MVVM.ObservableProperty[Nullable[bool]]]::new($false)
    }
}

class PrintTaskGroup {
    [string]      $Header
    [PrintTask[]] $Tasks

    PrintTaskGroup([string]$header, [PrintTask[]]$tasks) {
        $this.Header = $header
        $this.Tasks  = $tasks
    }
}

class PrintTaskModel {
    [PrintTaskGroup[]] $Groups

    PrintTaskModel() {
        [string]$JsonPath = Join-Path $PSScriptRoot "PrintTaskGroups.json"
        $raw = Get-Content $JsonPath -Raw | ConvertFrom-Json

        $this.Groups = foreach ($g in $raw) {
            $tasks = foreach ($t in $g.Tasks) {
                [PrintTask]::new($t.Name, $t.RequiresSheetPath, $t.ActionId)
            }

            [PrintTaskGroup]::new($g.Header, $tasks)
        }
    }

    [PrintTaskGroup[]] GetPrintTaskGroups() {
        return $this.Groups
    }
}
