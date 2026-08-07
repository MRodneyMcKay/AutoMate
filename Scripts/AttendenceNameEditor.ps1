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

$nameEditorPath = Join-Path $PSScriptRoot "..\Modules\WPFToolkit\Controls\NameEditor"

. (Join-Path $nameEditorPath "NameEditorViewModel.ps1")
. (Join-Path $nameEditorPath "NameEditorView.ps1")

$nameEditorPath = Join-Path $PSScriptRoot "..\Modules\WPFToolkit\Controls\InputDialog"

. (Join-Path $nameEditorPath "InputDialogField.ps1")
. (Join-Path $nameEditorPath "InputDialogViewModel.ps1")
. (Join-Path $nameEditorPath "InputDialogView.ps1")


. "$PSScriptRoot\Models\AttendenceModel.ps1"

. "$PSScriptRoot\ViewModels\AttendencePositionViewModel.ps1"
. "$PSScriptRoot\ViewModels\AttendenceDepartmentViewModel.ps1"
. "$PSScriptRoot\ViewModels\AttendenceNameEditorViewModel.ps1"

. "$PSScriptRoot\Views\AttendenceNameEditorView.ps1"


$app = [System.Windows.Application]::Current
if(-not $app){
    $app = [System.Windows.Application]::new()
    $app.ShutdownMode = "OnExplicitShutdown"
    $app.Resources.MergedDictionaries.Add((Get-ThemeResources))
    $app.Resources.MergedDictionaries.Add((Get-ThemeStyle))
}

$model =[AttendenceModel]::new((Join-Path $PSScriptRoot "Data\nevek.xml"))
$viewModel = [AttendenceNameEditorViewModel]::new($model)
$view = [AttendenceNameEditorView]::new($viewModel)
$frame = [System.Windows.Threading.DispatcherFrame]::new()
$view.Window.Add_Closed({ $frame.Continue = $false}.GetNewClosure())
$view.Show()

[System.Windows.Threading.Dispatcher]::PushFrame($frame)

Remove-Variable `
    app,
    model,
    viewModel,
    view,
    frame `
    -ErrorAction SilentlyContinue

[GC]::Collect()