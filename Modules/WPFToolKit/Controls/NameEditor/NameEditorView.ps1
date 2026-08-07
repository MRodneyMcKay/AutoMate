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

using namespace System.Windows
using namespace System.Windows.Controls
using namespace System.Windows.Markup

<#
    Loaded after NameEditorViewModel.ps1 - the SetViewModel() method parameter is
    typed [NameEditorViewModel], and PowerShell resolves class member types at
    parse time, so that class must already exist when this file is parsed.
#>
class NameEditorView {
    [FrameworkElement]  $Root
    [ListBox]           $NameList
    [TextBox]           $InputBox
    hidden [NameEditorViewModel] $ViewModel

    NameEditorView() {
        $xamlPath = Join-Path $PSScriptRoot 'NameEditorView.xaml'
        if (-not (Test-Path -Path $xamlPath)) {
            throw "NameEditor XAML not found: $xamlPath"
        }

        $xaml = Get-Content -Path $xamlPath -Raw
        $this.Root = [XamlReader]::Parse($xaml)
        $this.NameList = $this.Root.FindName('NameList')
        $this.InputBox = $this.Root.FindName('InputBox')
    }
    [void] SetViewModel([NameEditorViewModel] $viewModel) {

        $this.ViewModel = $viewModel

        $this.Root.DataContext = $viewModel


        # SelectedItemsSyncBehavior bridge
        $this.NameList.Tag = {
            param($items)

            $viewModel.SetSelection($items)

        }.GetNewClosure()


        # ClickEmptyListSpaceCancelBehavior bridge
        [WPFToolkit.Behaviors.BehaviorBridge]::SetAction(
            $this.NameList,
            {
                $viewModel.ExecuteCancel()
            }.GetNewClosure()
        )
        $nl = $this.NameList
        $this.viewModel.RequestUnselectAll = {
            $nl.UnselectAll()
        }.GetNewClosure()
    }
}