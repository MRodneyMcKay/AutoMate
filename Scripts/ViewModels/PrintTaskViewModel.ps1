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

class PrintTaskViewModel : WPFToolkit.MVVM.ObservableObject {
    [PrintTaskModel] $Model
    [PrintTaskService] $PrintService

    [System.Collections.ObjectModel.ObservableCollection[PrintTaskGroup]] $Groups
    [System.Collections.ObjectModel.ObservableCollection[PrintTask]] $SelectedTasks

    [WPFToolkit.MVVM.ObservableProperty[string]] $SelectedSheetPath
    [System.Windows.Window] $Window

    [WPFToolkit.MVVM.ObservableProperty[Nullable[bool]]] $SelectAll
    hidden [bool] $_SuppressSelectAllCascade

    [WPFToolkit.MVVM.RelayCommand] $ShowNamesCommand
    [WPFToolkit.MVVM.RelayCommand] $InvokeSelectedPrintTasksCommand
    [WPFToolkit.MVVM.RelayCommand] $BrowseSheetCommand
    [WPFToolkit.MVVM.RelayCommand] $CloseCommand


    PrintTaskViewModel() {

        $this.Model = [PrintTaskModel]::new()
        $this.PrintService = [PrintTaskService]::new()
        $this.SelectedSheetPath = [WPFToolkit.MVVM.ObservableProperty[string]]::new()

        $this.Groups = [System.Collections.ObjectModel.ObservableCollection[PrintTaskGroup]]::new()

        foreach ($group in $this.Model.GetPrintTaskGroups()) {
            $this.Groups.Add($group)
        }
        $vm = $this
        foreach ($task in $this.GetAllTasks()) {
            $task.IsSelected.add_PropertyChanged({
                    $vm.RecomputeSelectAllState()
                    [System.Windows.Input.CommandManager]::InvalidateRequerySuggested()
                }.GetNewClosure())
        }

        $this.SelectAll = [WPFToolkit.MVVM.ObservableProperty[Nullable[bool]]]::new($false)
        $this.SelectAll.add_PropertyChanged({
                if ($vm._SuppressSelectAllCascade) { return }

                $target = ($vm.SelectAll.Value -eq $true)
                foreach ($task in $vm.GetAllTasks()) {
                    $task.IsSelected.Value = $target
                }
            }.GetNewClosure())

        $this.InvokeSelectedPrintTasksCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($vm)
                $vm.InvokeSelectedPrintTasks()
            },
            [System.Predicate[object]] {
                param($vm)

                if ($null -eq $vm) { return $false }
                if ($null -eq $vm.Groups) { return $false }

                $selectedTasks = $vm.GetSelectedTasks()

                # Nothing selected
                if ($selectedTasks.Count -eq 0) {
                    return $false
                }

                # A selected task requires an Excel sheet, but none is chosen
                $requiresSheet = $selectedTasks | Where-Object {
                    $_.RequiresSheetPath -eq $true
                }

                if ($requiresSheet -and [string]::IsNullOrWhiteSpace($vm.SelectedSheetPath.Value)) {
                    return $false
                }

                return $true
            }
        )

        $this.BrowseSheetCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($vm)
                $vm.BrowseSheet()
            }
        )

        $this.ShowNamesCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($vm)
                $vm.ShowNames()
            }
        )

        $this.CloseCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($vm)
                $vm.CloseWindow()
            }
        )
    }

    [PrintTask[]] GetAllTasks() {
        return @(
            foreach ($group in $this.Groups) {
                $group.Tasks
            }
        )
    }

    [void] RecomputeSelectAllState() {
        $allTasks = $this.GetAllTasks()

        if ($allTasks.Count -eq 0) {
            $newState = $false
        }
        else {
            $selectedCount = @($allTasks | Where-Object { $_.IsSelected.Value -eq $true }).Count

            if ($selectedCount -eq 0) { $newState = $false }
            elseif ($selectedCount -eq $allTasks.Count) { $newState = $true }
            else { $newState = $null }
        }

        $this._SuppressSelectAllCascade = $true
        try {
            $this.SelectAll.Value = $newState
        }
        finally {
            $this._SuppressSelectAllCascade = $false
        }
    }

    [PrintTask[]] GetSelectedTasks() {
        return @($this.GetAllTasks() | Where-Object { $_.IsSelected.Value -eq $true })
    }

    [void] CloseWindow() {
        if ($this.Window -ne $null) {
            $this.Window.Close()
        }
    }

    [void] BrowseSheet() {
        $dialog = [Microsoft.Win32.OpenFileDialog]::new()
        $dialog.Title = "Select Excel sheet"
        $dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"

        if ($dialog.ShowDialog() -eq $true) {
            $this.SetSelectedSheet($dialog.FileName)
        }
    }

    [void] SetSelectedSheet([string]$path) {
        if ($this.ValidateSheet($path)) {
            $this.SelectedSheetPath.Value = $path
        }

        [System.Windows.Input.CommandManager]::InvalidateRequerySuggested()
    }

    [bool] ValidateSheet([string]$path) {

        if ([string]::IsNullOrWhiteSpace($path)) {
            return $false
        }
        if (-not (Test-Path $path)) {
            return $false
        }

        [xml]$xml = $null

        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
            try {
                $entry = $zip.GetEntry("xl/workbook.xml")
                if ($null -eq $entry) {
                    return $false
                }
                $reader = [System.IO.StreamReader]::new($entry.Open())
                try {
                    $xml = [xml]$reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $zip.Dispose()
            }
        }
        catch {
            return $false
        }
        if ($null -eq $xml) {
            return $false
        }
        $sheets = $xml.SelectNodes(
            "//*[local-name()='sheet']"
        )
        foreach ($sheet in $sheets) {
            if ($sheet.name -eq "Fizikai") {
                return $true
            }
        }

        return $false
    }

    [void] InvokeSelectedPrintTasks() {

        $this.SelectedTasks = $this.GetSelectedTasks()
        foreach ($task in $this.SelectedTasks) {
            $this.PrintService.Execute(
                $task,
                $this.SelectedSheetPath.Value
            )
        }
    }

    [void] ShowNames () {
        $scriptPath = Join-Path `
        $PSScriptRoot `
        "..\AttendenceNameEditor.ps1"

    $scriptPath = [System.IO.Path]::GetFullPath($scriptPath)


    if (-not (Test-Path $scriptPath)) {
        throw "Name editor script not found: $scriptPath"
    }


    Start-Process `
        -FilePath "pwsh.exe" `
        -ArgumentList @(
            "-NoProfile",
            "-File",
            "`"$scriptPath`""
        ) `
        -WindowStyle Hidden
    }
}