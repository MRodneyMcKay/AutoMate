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

class WpfView {
    [string] $XamlPath
    [System.Windows.Window] $Window
    WpfView() {
        $this.XamlPath = Join-Path $PSScriptRoot "PrintManagerView.xaml"
        $xaml = Get-Content $this.XamlPath -Raw
        $this.Window = [Windows.Markup.XamlReader]::Parse($xaml)
        $this.SetupDragDrop()
    }

    [void] SetDataContext([object] $context) {
        $this.Window.DataContext = $context
        $context.Window = $this.Window
    }

    hidden [void] SetupDragDrop() {
        $view = $this
        $dropArea = $this.Window.FindName("SheetDropArea")
        $dropArea.AllowDrop = $true
        $normalBackground = $dropArea.Background

        $dropArea.Add_DragOver({
                param($sender, $e)
                $allowed = $false
                if ($e.Data.GetDataPresent(
                        [System.Windows.DataFormats]::FileDrop
                    )) {
                    $files = $e.Data.GetData(
                        [System.Windows.DataFormats]::FileDrop
                    )
                    if ($files.Count -eq 1) {
                        $ext = [System.IO.Path]::GetExtension($files[0])
                        if ($ext -ieq ".xlsx") {
                            $allowed = $true
                        }
                    }
                }
                if ($allowed) {
                    $e.Effects = [System.Windows.DragDropEffects]::Copy
                }
                else {
                    $e.Effects = [System.Windows.DragDropEffects]::None
                }
                $e.Handled = $true
            }.GetNewClosure())

        $dropArea.Add_DragEnter({
                param($sender, $e)
                if ($e.Data.GetDataPresent(
                        [System.Windows.DataFormats]::FileDrop
                    )) {
                    $sender.Background =
                    $view.Window.FindResource("DropHighlightBrush")
                    $e.Effects = [System.Windows.DragDropEffects]::Copy
                }
                $e.Handled = $true
            }.GetNewClosure())

        $dropArea.Add_DragLeave({
                param($sender, $e)
                $sender.Background = $normalBackground
            }.GetNewClosure())

        $dropArea.Add_Drop({
                param($sender, $e)
                $sender.Background = $normalBackground
                if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
                    return
                }
                $files = $e.Data.GetData(
                    [System.Windows.DataFormats]::FileDrop
                )
                if ($files.Count -eq 0) {
                    return
                }
                $ext = [System.IO.Path]::GetExtension($files[0])
                if ($ext -ine ".xlsx") {
                    return
                }
                $view.Window.DataContext.SetSelectedSheet($files[0])
            }.GetNewClosure())
    }

    [void] Show() {
        $this.Window.ShowDialog()
    }

    [void] Close() {
        $this.Window.Close()
    }
}
