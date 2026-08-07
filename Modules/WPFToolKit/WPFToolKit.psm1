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

$DllPath = Join-Path $PSScriptRoot 'Bin\WPFToolkit.dll'
Add-Type -Path $DllPath

. "$PSScriptRoot\ColorMath.ps1"
. "$PSScriptRoot\ThemeManager.ps1"
. "$PSScriptRoot\BehaviorRegistry.ps1"

Initialize-BehaviorBridge

Get-ChildItem "$PSScriptRoot\Controls\NameEditor\Behaviors\*.ps1" |
    ForEach-Object {
        . $_.FullName
    }

function Get-WpfApplication {
    return [WpfApplicationHost]::Current
}
Export-ModuleMember -Function Get-ThemeResources, Get-ThemeStyle, Get-WpfApplication, Initialize-BehaviorBridge, Register-Behavior