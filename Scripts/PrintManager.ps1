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

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

Import-Module (Join-Path $PSScriptRoot "..\Modules\WPFToolkit\WPFToolkit.psd1") -Force
Import-Module (Join-Path $PSScriptRoot "..\Modules\LoggingSystem\LoggingSystem.psd1") -Force
Import-Module (Join-Path $PSScriptRoot "..\Modules\PrintNextMonthFolder\PrintNextMonthFolder.psd1") -Force

. "$PSScriptRoot\Models\PrintTaskModel.ps1"
. "$PSScriptRoot\Services\PrintTaskService.ps1"
. "$PSScriptRoot\ViewModels\PrintTaskViewModel.ps1"
. "$PSScriptRoot\Views\PrintManagerView.ps1"

$app = [System.Windows.Application]::Current
if (-not $app) {
    $app = [System.Windows.Application]::new()
    $app.ShutdownMode = 'OnExplicitShutdown'
    $app.Resources.MergedDictionaries.Add((Get-ThemeResources))
    $app.Resources.MergedDictionaries.Add((Get-ThemeStyle))
}
$vm = [PrintTaskViewModel]::new()
$view = [WpfView]::new()
$view.SetDataContext($vm)
$frame = New-Object System.Windows.Threading.DispatcherFrame
$view.window.Add_Closed({ $frame.Continue = $false })
$view.window.Show() 
[System.Windows.Threading.Dispatcher]::PushFrame($frame)

Remove-Variable window, frame -ErrorAction SilentlyContinue
[System.GC]::Collect()