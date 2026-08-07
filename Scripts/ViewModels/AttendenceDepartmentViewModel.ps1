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

class AttendenceDepartmentViewModel : WPFToolkit.MVVM.ObservableObject {
    [AttendenceDepartment] $Model
    [WPFToolkit.MVVM.ObservableProperty[bool]] $IsDirty
    [System.Collections.ObjectModel.ObservableCollection[AttendencePositionViewModel]] $Positions
    [WPFToolkit.MVVM.ObservableProperty[string]] $DisplayName
    [WPFToolkit.MVVM.ObservableProperty[AttendencePositionViewModel]] $SelectedPosition

    AttendenceDepartmentViewModel( [AttendenceDepartment] $model ){
        $DepartmentVM = $this
        $this.Model = $model
        $this.IsDirty = [WPFToolkit.MVVM.ObservableProperty[bool]]::new( $false )
        $this.DisplayName = [WPFToolkit.MVVM.ObservableProperty[string]]::new( "" )
        $this.SelectedPosition = [WPFToolkit.MVVM.ObservableProperty[AttendencePositionViewModel]]::new( $null )
        $this.IsDirty.add_PropertyChanged({
            $DepartmentVM.RefreshDisplayName()
        }.GetNewClosure())
        $this.Positions = [System.Collections.ObjectModel.ObservableCollection[AttendencePositionViewModel]]::new()
        foreach($position in $model.Positions){
            $this.Positions.Add([AttendencePositionViewModel]::new( $position, $this ))
        }
        if($this.Positions.Count -gt 0){
            $this.SelectedPosition.Value = $this.Positions[0]
        }
        $this.RefreshDisplayName()
    }
    [void] RefreshDisplayName(){
        $facility = ""
        if(-not [string]::IsNullOrWhiteSpace( $this.Model.Facility)  -and   $this.Model.Facility -ne "Kecskeméti Fürdő" ){
            $facility = " - $($this.Model.Facility)"
        }
        $prefix = ""
        if($this.IsDirty.Value){
            $prefix = "● "
        }
        $this.DisplayName.Value = "$prefix$($this.Model.Name)$facility"
    }
    [void] MarkDirty(){
        $this.IsDirty.Value = $true
    }
    [void] ClearDirty(){
        $this.IsDirty.Value = $false
    }
    [void] AddPosition([AttendencePosition] $position ){
        $positionVM = [AttendencePositionViewModel]::new( $position, $this )
        $this.Positions.Add( $positionVM )
        $this.SelectedPosition.Value = $positionVM
    }
}