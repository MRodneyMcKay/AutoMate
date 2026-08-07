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
using namespace System.Windows.Controls

class AttendenceNameEditorView {
    [string] $XamlPath
    [Window] $Window
    [AttendenceNameEditorViewModel] $ViewModel
    [NameEditorView] $CurrentNameEditor

    AttendenceNameEditorView( [AttendenceNameEditorViewModel] $viewModel  ){
        $this.ViewModel = $viewModel
        $this.XamlPath = Join-Path $PSScriptRoot "AttendenceNameEditorView.xaml"
        $xaml = Get-Content $this.XamlPath -Raw
        $this.Window = [Window][XamlReader]::Parse( $xaml )
        $this.Window.DataContext = $this.ViewModel
        $this.AttachNameEditor()
        $this.ViewModel.Window = $this.Window
    }

    [void] AttachNameEditor(){
        $contentHost = $this.Window.FindName("NameEditorContent")
        if($null -eq $contentHost){
            return
        }
        $view = $this

        $this.ViewModel.SelectedDepartment.add_PropertyChanged({

            $view.AttachPositionListener(
                $contentHost
            )
            $view.RefreshNameEditor(
                $contentHost
            )
        }.GetNewClosure())

        $this.AttachPositionListener(
            $contentHost
        )

        $this.RefreshNameEditor(
            $contentHost
        )
    }

    [void] AttachPositionListener([ContentControl] $contentHost){
        $department = $this.ViewModel.SelectedDepartment.Value
        if($null -eq $department){
            return
        }

        $view = $this

        #
        # munkakör tab váltás
        #
        $department.SelectedPosition.add_PropertyChanged({
            $view.RefreshNameEditor(
                $contentHost
            )
        }.GetNewClosure())
    }

    [void] RefreshNameEditor( [ContentControl] $contentHost ){
        $department = $this.ViewModel.SelectedDepartment.Value
        if($null -eq $department){
            $contentHost.Content = $null
            return
        }

        $position = $department.SelectedPosition.Value
        if($null -eq $position){
            $contentHost.Content = $null
            return
        }
        $editor = [NameEditorView]::new()
        $editorVM = $this.ViewModel.CreateNameEditorViewModel( $position )
        $editor.SetViewModel( $editorVM)
        $this.CurrentNameEditor = $editor
        $contentHost.Content = $editor.Root
    }

    [void] Show(){
        $this.Window.ShowDialog()
    }

    [void] Close() {
        $this.Window.Close()
    }
}