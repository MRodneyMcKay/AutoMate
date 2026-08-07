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

$script:Behaviors = @{}

function Register-Behavior {
    param(
        [string]$Name,
        [scriptblock]$Attach
    )
    $script:Behaviors[$Name] = $Attach
}
function Initialize-BehaviorBridge {
    [WPFToolkit.Behaviors.BehaviorBridge]::add_BehaviorRequested({
        param($sender,$event)
        try {
            $behavior = $script:Behaviors[$event.BehaviorName]
            if ($null -ne $behavior) {
                & $behavior $event.Element
            }
        }
        catch {
            Write-Host "BEHAVIOR FAILED: $($event.BehaviorName)"
            Write-Host $_.Exception.Message
            Write-Host $_.ScriptStackTrace
            throw
        }
    })
}