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

class AttendenceNameEditorViewModel : WPFToolkit.MVVM.ObservableObject {
    [AttendenceModel] $Model
    [System.Windows.Window] $Window
    [System.Collections.ObjectModel.ObservableCollection[AttendenceDepartmentViewModel]] $Departments
    [WPFToolkit.MVVM.ObservableProperty[AttendenceDepartmentViewModel]] $SelectedDepartment
    [WPFToolkit.MVVM.RelayCommand] $CloseCommand
    [WPFToolkit.MVVM.RelayCommand] $SaveCommand
    [WPFToolkit.MVVM.RelayCommand] $ReloadCommand
    [WPFToolkit.MVVM.RelayCommand] $AddDepartmentCommand
    [WPFToolkit.MVVM.RelayCommand] $AddPositionCommand
    [WPFToolkit.MVVM.ObservableProperty[bool]] $IsDirty
    [WPFToolkit.MVVM.ObservableProperty[string]] $StatusText

    AttendenceNameEditorViewModel(
        [AttendenceModel] $model
    ) {
        $this.Model = $model
        $vm = $this
        $this.IsDirty = [WPFToolkit.MVVM.ObservableProperty[bool]]::new($false)
        $this.StatusText = [WPFToolkit.MVVM.ObservableProperty[string]]::new( "Készen áll" )
        $this.Departments = [System.Collections.ObjectModel.ObservableCollection[AttendenceDepartmentViewModel]]::new()
        $this.SelectedDepartment = [WPFToolkit.MVVM.ObservableProperty[AttendenceDepartmentViewModel]]::new( $null)
        foreach ($department in $this.Model.Departments) {
            $departmentVM = [AttendenceDepartmentViewModel]::new( $department )
            $this.RegisterDepartment( $departmentVM )
            $this.Departments.Add( $departmentVM )
        }
        if ($this.Departments.Count -gt 0) {
            $this.SelectedDepartment.Value = $this.Departments[0]
        }
        $this.CloseCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($parameter)
                $vm.CloseWindow()
            }.GetNewClosure(),
            {
                return -not $vm.IsDirty.Value
            }.GetNewClosure()
        )
        $this.AddDepartmentCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($parameter)
                $vm.AddDepartment()
            }.GetNewClosure()
        )
        $this.AddPositionCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($department)
                if ($department -is [AttendenceDepartmentViewModel]) {
                    $vm.AddPosition( $department )
                }
            }.GetNewClosure()
        )
        $this.ReloadCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($parameter)
                $vm.Reload()
            }.GetNewClosure(),
            {
                return $vm.IsDirty.Value
            }.GetNewClosure()
        )
        $this.SaveCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]] {
                param($parameter)
                $vm.Save()
            }.GetNewClosure(),
            {
                return $vm.IsDirty.Value
            }.GetNewClosure()
        )
    }
    [void] RegisterDepartment( [AttendenceDepartmentViewModel] $departmentVM) {
        $vm = $this
        $departmentVM.IsDirty.add_PropertyChanged({
            if($departmentVM.IsDirty.Value){
                $vm.IsDirty.Value = $true
            }
        }.GetNewClosure())
    }
    [void] AddDepartment() {
        $dialogVM = [InputDialogViewModel]::new()
        $dialogVM.Title.Value = "Új részleg"
        $dialogVM.Fields.Add( [InputDialogField]::new(  "Név" ) )
        $dialogVM.Fields.Add( [InputDialogField]::new("Létesítmény" ) )
        $dialog = [InputDialogView]::new( $dialogVM)
        if (-not $dialog.Show()) {
            return
        }
        $values = $dialogVM.GetValues()
        if ([string]::IsNullOrWhiteSpace( $values["Név"] ) ) {
            return
        }
        $department = [AttendenceDepartment]::new()
        $department.Name = $values["Név"]
        $department.Facility = $values["Létesítmény"]
        $this.Model.Departments.Add( $department)
        $departmentVM = [AttendenceDepartmentViewModel]::new( $department )
        $this.RegisterDepartment( $departmentVM)
        $departmentVM.MarkDirty()
        $this.Departments.Add( $departmentVM )
        $this.SelectedDepartment.Value = $departmentVM
        $this.AddPosition(  $departmentVM  )
        $this.SetStatus("$($values['Név']) hozzáadva")
    }

    [void] AddPosition( [AttendenceDepartmentViewModel] $departmentVM ) {
        if ($null -eq $departmentVM) {
            return
        }
        $dialogVM = [InputDialogViewModel]::new()
        $dialogVM.Title.Value = "Új munkakör"
        $dialogVM.Fields.Add( [InputDialogField]::new(  "Név" ) )
        $dialog = [InputDialogView]::new( $dialogVM )
        if (-not $dialog.Show()) {
            return
        }
        $values = $dialogVM.GetValues()
        if (  [string]::IsNullOrWhiteSpace( $values["Név"] ) ) {
            return
        }
        $position = [AttendencePosition]::new()
        $position.Name = $values["Név"]
        $departmentVM.Model.Positions.Add($position)
        $positionVM = [AttendencePositionViewModel]::new( $position, $departmentVM)
        $positionVM.MarkDirty()
        $departmentVM.Positions.Add( $positionVM )
        $departmentVM.SelectedPosition.Value = $positionVM
        $this.SetStatus("$($values['Név']) munkakör hozzáadva")
    }
    
    [NameEditorViewModel] CreateNameEditorViewModel( [AttendencePositionViewModel] $positionVM ) {
        $vmodel = $this
        $editorVM = [NameEditorViewModel]::new( $positionVM.Model.Names )
        $editorVM.OnAdd = {
            param( [string] $name)
            $positionVM.AddName( $name )
            $vmodel.SetStatus( "Név hozzáadva: $name" )
        }.GetNewClosure()
        $editorVM.OnUpdate = {
            param(
                [string] $oldName,
                [string] $newName
            )
            $positionVM.UpdateName(  $oldName, $newName )
            $vmodel.SetStatus("Név módosítva: $newName")
        }.GetNewClosure()
        $editorVM.OnDelete = {
            param( [object[]] $names )
            $positionVM.RemoveNames( $names)
            $vmodel.SetStatus( "$($names.Count) név törölve")
        }.GetNewClosure()
        return $editorVM
    }
    [void] CloseWindow() {
        if ($null -ne $this.Window) {
            $this.Window.Close()
        }
    }
    [void] Reload() {
        $this.Model.Departments.Clear()
        $this.Model.Load()
        $this.Departments.Clear()
        foreach ($department in $this.Model.Departments) {
            $departmentVM = [AttendenceDepartmentViewModel]::new( $department )
            $this.RegisterDepartment( $departmentVM )
            $this.Departments.Add( $departmentVM )
        }

        if ($this.Departments.Count -gt 0) {
            $this.SelectedDepartment.Value = $this.Departments[0]
        } else {
            $this.SelectedDepartment.Value = $null
        }

        foreach ($department in $this.Departments) {
            $department.ClearDirty()
            foreach ($position in $department.Positions) {
                $position.ClearDirty()
            }
        }
        $this.IsDirty.Value = $false
        $this.SetStatus("Módosítások elvetve")
    }
    [void] Save() {
        $this.SetStatus("Módosítások mentése")
        $this.Model.Save()
        foreach ($department in $this.Departments) {
            $department.ClearDirty()
            foreach ($position in $department.Positions) {
                $position.ClearDirty()
            }
        }
        $this.IsDirty.Value = $false
        $this.SetStatus("Módosítások mentve")
    }
    [void] SetStatus( [string] $message ) {
        $this.StatusText.Value = $message
    }
}