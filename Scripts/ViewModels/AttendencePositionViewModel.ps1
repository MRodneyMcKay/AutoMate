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

class AttendencePositionViewModel : WPFToolkit.MVVM.ObservableObject {
    [AttendencePosition] $Model
    [object] $Department
    [WPFToolkit.MVVM.ObservableProperty[bool]] $IsDirty
    [WPFToolkit.MVVM.ObservableProperty[string]] $DisplayName

    AttendencePositionViewModel(
        [AttendencePosition] $model,
        [object] $department
    ){
        $PositionVM = $this
        $this.Model = $model
        $this.Department = $department
        $this.IsDirty = [WPFToolkit.MVVM.ObservableProperty[bool]]::new( $false )
        $this.DisplayName = [WPFToolkit.MVVM.ObservableProperty[string]]::new( "" )
        $this.IsDirty.add_PropertyChanged({
            $PositionVM.RefreshDisplayName()
        }.GetNewClosure())
        $this.RefreshDisplayName()
    }



    [void] RefreshDisplayName(){
        $this.DisplayName.Value = "$($this.IsDirty.Value ? '● ' : '')$($this.Model.Name)"
    }

    [void] AddName([string] $name){
        $name = $name.Trim()
        if( [string]::IsNullOrWhiteSpace( $name ) ){
            return
        }

        $exists =  $this.Model.Names | Where-Object { $_.Equals( $name, [System.StringComparison]::InvariantCultureIgnoreCase ) }
        if(-not $exists){
            $this.Model.Names.Add( $name )
            $this.MarkDirty()
        }
    }

    [void] UpdateName(
        [string] $oldName,
        [string] $newName
    ){
        $newName = $newName.Trim()
        if( [string]::IsNullOrWhiteSpace( $newName )){
            return
        }
        $index = $this.Model.Names.IndexOf( $oldName )
        if($index -ge 0){
            $this.Model.Names[$index] = $newName
            $this.MarkDirty()
        }
    }

    [void] RemoveNames( [object[]] $names ){
        $changed = $false
        foreach($name in $names){
            if( $this.Model.Names.Remove( $name )){
                $changed = $true
            }
        }
        if($changed){
            $this.MarkDirty()
        }
    }

    [void] MarkDirty(){
        $this.IsDirty.Value = $true
        if($this.Department){
            $this.Department.MarkDirty()
        }
    }

    [void] ClearDirty(){
        $this.IsDirty.Value = $false
    }
}