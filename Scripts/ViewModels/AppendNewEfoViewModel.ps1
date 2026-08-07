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

class AppendNewEfoViewModel : WPFToolkit.MVVM.ObservableObject {
    [EfoNamesModel] $Model
    [System.Windows.Window] $Window
    [WPFToolkit.MVVM.ObservableProperty[string]]  $StatusText
    [WPFToolkit.MVVM.ObservableProperty[bool]] $IsDirty
    [WPFToolkit.MVVM.RelayCommand] $SaveCommand
    [WPFToolkit.MVVM.RelayCommand] $ReloadCommand
    [WPFToolkit.MVVM.RelayCommand] $CloseCommand

    AppendNewEfoViewModel(
        [EfoNamesModel] $model
    ) {
        $this.Model = $model
        $this.StatusText = [WPFToolkit.MVVM.ObservableProperty[string]]::new("Készen áll")
        $this.IsDirty = [WPFToolkit.MVVM.ObservableProperty[bool]]::new($false)
        $this.SaveCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [Action[object]] {
                param($p)
                $vm.Save()
            },
            {
                return $vm.IsDirty.Value
            }.GetNewClosure()
        )
        $this.ReloadCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [Action[object]] {
                param($p)
                $vm.Reload()
            },
            {
                return $vm.IsDirty.Value
            }.GetNewClosure()
        )
        $this.CloseCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [Action[object]] {
                param($p)
                $vm.Close()
            },
            {
                return -not $vm.IsDirty.Value
            }.GetNewClosure()
        )
    }

    [void] AddName( [string] $name ) {
        $name = $name.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            return
        }
        if ( $this.Model.Names | Where-Object { $_.Equals( $name, [System.StringComparison]::InvariantCultureIgnoreCase ) } ) {
            $this.SetStatus("Duplikált név")
            return
        }
        $this.Model.Names.Add($name)
        $this.Model.Sort()
        $this.MarkDirty()
        $this.SetStatus( "Név hozzáadva: $name" )
    }

    [void] UpdateName(
        [string] $oldName,
        [string] $newName
    ) {
        $newName = $newName.Trim()
        if ([string]::IsNullOrWhiteSpace($newName)) {
            $this.SetStatus("A név nem lehet üres")
            return
        }
        $index = $this.Model.Names.IndexOf($oldName)
        if ($index -lt 0) {
            $this.SetStatus("A név nem található")
            return
        }
        $this.Model.Names[$index] = $newName
        $this.Model.Sort()
        $this.MarkDirty()
        $this.SetStatus("Név módosítva: $newName")
    }

    [void] DeleteNames(
        [object[]] $names
    ) {
        $count = 0
        foreach ($name in $names) {
            if ( $this.Model.Names.Remove( $name) ) {
                $count++
            }
        }
        if ($count -gt 0) {
            $this.MarkDirty()
            $this.SetStatus( "$count név törölve")
        }
    }

    [void] Save() {
        if ( $this.Model.Save() ) {
            $this.IsDirty.Value =  $false
            $this.SetStatus( "Mentés kész" )
        } else {
            $this.SetStatus( "Mentés sikertelen")
        }
    }

    [void] Reload() {
        $this.Model.Reload()
        $this.IsDirty.Value = $false
        $this.SetStatus("Módosítások elvetve")
    }

    [void] MarkDirty() {
        $this.IsDirty.Value =$true
    }

    [void] SetStatus(
        [string] $message
    ) {
        $this.StatusText.Value = $message
    }
    
    [void] Close() {
        if ($this.Window) {
            $this.Window.Close()
        }
    }
    $vm = $this
}