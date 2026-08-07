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
using namespace System.Windows.Markup
class AppendNewEfoView {
    [string] $XamlPath
    [Window] $Window
    [NameEditorView] $NameEditor

    AppendNewEfoView() {
        $this.XamlPath = Join-Path $PSScriptRoot "AppendNewEfoView.xaml"
        if(-not (Test-Path $this.XamlPath)) {
            throw "XAML not found: $($this.XamlPath)"
        }
        $xaml = Get-Content $this.XamlPath -Raw
        $this.Window = [XamlReader]::Parse($xaml)
        $this.NameEditor = [NameEditorView]::new()
        $editorHost = $this.Window.FindName("NameEditorHost")
        if($null -eq $editorHost) {
            throw "NameEditorHost not found in XAML"
        }
        $editorHost.Content = $this.NameEditor.Root
    }

    [void] SetViewModel( [AppendNewEfoViewModel] $viewModel ) {
        $this.Window.DataContext = $viewModel
        $viewModel.Window = $this.Window
        $editorVM =  [NameEditorViewModel]::new($viewModel.Model.Names)

        $editorVM.OnAdd = {
            param($name)
            $viewModel.AddName(
                $name
            )
        }.GetNewClosure()

        $editorVM.OnUpdate = {
            param(
                $oldName,
                $newName
            )
            $viewModel.UpdateName(
                $oldName,
                $newName
            )
        }.GetNewClosure()

        $editorVM.OnDelete = {
            param($names)
            $viewModel.DeleteNames(
                $names
            )
        }.GetNewClosure()

        $this.NameEditor.SetViewModel($editorVM)
    }

    [void] Show() {
        $this.Window.ShowDialog()
    }
    
    [void] Close() {
        $this.Window.Close()
    }
}