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

using namespace System.Collections.ObjectModel
using namespace System.Collections.Generic

class NameEditorViewModel : WPFToolkit.MVVM.ObservableObject {

    [ObservableCollection[string]] $Names
    [WPFToolkit.MVVM.ObservableProperty[string]]   $InputText

    [bool]   $ShowEditor     = $true
    [bool]   $ShowDelete     = $false
    [bool]   $HasInputText   = $false
    [string] $AddUpdateLabel = 'Hozzáad'

    [List[string]] $SelectedNames

    [WPFToolkit.MVVM.RelayCommand] $AddUpdateCommand
    [WPFToolkit.MVVM.RelayCommand] $DeleteCommand
    [WPFToolkit.MVVM.RelayCommand] $CancelCommand

    # Callbacks supplied by the consumer. Always invoked with explicit parameters,
    # never relying on the caller's own captured/closed-over variables.
    [scriptblock] $OnAdd
    [scriptblock] $OnUpdate
    [scriptblock] $OnDelete
    [scriptblock] $OnCancel
    [scriptblock] $OnSelectionChanged

    # Wired by Register-NameEditorBehavior to the real ListBox's UnselectAll().
    # This ViewModel has no reference to the control, so it asks the View to
    # clear the actual selection rather than only clearing its own bookkeeping.
    [scriptblock] $RequestUnselectAll

    NameEditorViewModel() {
        $this.Initialize([ObservableCollection[string]]::new())
    }

    NameEditorViewModel([ObservableCollection[string]] $names) {
        $this.Initialize($names)
    }

    hidden [void] Initialize([ObservableCollection[string]] $names) {
        $this.Names         = $names
        $this.InputText     = [WPFToolkit.MVVM.ObservableProperty[string]]::new('')
        $this.SelectedNames = [List[string]]::new()

        $vm = $this

        $this.InputText.add_PropertyChanged({
            param($Sender, $Event)
            $vm.HasInputText = -not [string]::IsNullOrWhiteSpace($vm.InputText.Value)
            $vm.OnPropertyChanged('HasInputText')
        }.GetNewClosure())

        $this.AddUpdateCommand = [WPFToolkit.MVVM.RelayCommand]::new(
            [Action[object]]({ param($p) $vm.ExecuteAddUpdate() }.GetNewClosure()),
            [Predicate[object]]({ param($p) -not [string]::IsNullOrWhiteSpace($vm.InputText.Value) }.GetNewClosure())
        )
        $this.DeleteCommand = [WPFToolkit.MVVM.RelayCommand]::new(
            [Action[object]]({ param($p) $vm.ExecuteDelete() }.GetNewClosure()),
            [Predicate[object]]({ param($p) $vm.SelectedNames.Count -gt 0 }.GetNewClosure())
        )
        $this.CancelCommand = [WPFToolkit.MVVM.RelayCommand]::new(
            [Action[object]]({ param($p) $vm.ExecuteCancel() }.GetNewClosure())
        )
    }


    # Called by the View's ListBox.SelectionChanged handler - SelectedItems isn't
    # a bindable DP, so this is the one place selection crosses the View/ViewModel
    # boundary explicitly instead of through a binding.
    [void] SetSelection([object[]] $selectedItems) {
        $this.SelectedNames.Clear()
        foreach ($item in $selectedItems) { [void]$this.SelectedNames.Add([string]$item) }

        switch ($this.SelectedNames.Count) {
            0 {
                $this.InputText.Value = ''
                $this.ShowEditor      = $true
                $this.AddUpdateLabel  = 'Hozzáad'
                $this.ShowDelete      = $false
            }
            1 {
                $this.InputText.Value = $this.SelectedNames[0]
                $this.ShowEditor      = $true
                $this.AddUpdateLabel  = 'Frissítés'
                $this.ShowDelete      = $true
            }
            default {
                $this.InputText.Value = ''
                $this.ShowEditor      = $false
                $this.ShowDelete      = $true
            }
        }

        $this.OnPropertyChanged('ShowEditor')
        $this.OnPropertyChanged('ShowDelete')
        $this.OnPropertyChanged('AddUpdateLabel')

        if ($this.OnSelectionChanged) {
            & $this.OnSelectionChanged @($this.SelectedNames)
        }
    }

    [void] ExecuteAddUpdate() {
        $text = $this.InputText.Value.Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return
        }

        if ($this.SelectedNames.Count -eq 1) {
            if ($this.OnUpdate) {
                & $this.OnUpdate $this.SelectedNames[0] $text
            }
        }
        else {
            $lines = $text -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            foreach ($line in $lines) {
                if ($this.OnAdd) {
                    & $this.OnAdd $line
                }
            }
        }
        $this.ExecuteCancel()
    }

    [void] ExecuteDelete() {
        if ($this.SelectedNames.Count -eq 0) { return }
        if ($this.OnDelete) { & $this.OnDelete @($this.SelectedNames) }
    }

    [void] ExecuteCancel() {
        $this.InputText.Value = ''
        $this.SelectedNames.Clear()
        $this.ShowEditor      = $true
        $this.ShowDelete      = $false
        $this.AddUpdateLabel  = 'Hozzáad'

        $this.OnPropertyChanged('ShowEditor')
        $this.OnPropertyChanged('ShowDelete')
        $this.OnPropertyChanged('AddUpdateLabel')

        # Bring the real ListBox selection back in line with the state above -
        # SelectedNames is only this ViewModel's copy, not the actual control.
        if ($this.RequestUnselectAll) {
            & $this.RequestUnselectAll
        }

        if ($this.OnCancel) { & $this.OnCancel }
    }
}

