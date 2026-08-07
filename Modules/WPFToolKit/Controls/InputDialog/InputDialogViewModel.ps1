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


class InputDialogViewModel :
    WPFToolkit.MVVM.ObservableObject {


    [WPFToolkit.MVVM.ObservableProperty[string]]
    $Title


    [ObservableCollection[InputDialogField]]
    $Fields


    [WPFToolkit.MVVM.RelayCommand]
    $OkCommand


    [WPFToolkit.MVVM.RelayCommand]
    $CancelCommand


    [scriptblock]
    $CloseAction



    InputDialogViewModel(){


        $this.Title =
            [WPFToolkit.MVVM.ObservableProperty[string]]::new(
                ""
            )


        $this.Fields =
            [ObservableCollection[InputDialogField]]::new()



        $vm = $this



        $this.OkCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]]{
                
                if($null -ne $vm.CloseAction){

                    & $vm.CloseAction $true

                }

            }.GetNewClosure()
        )



        $this.CancelCommand =
        [WPFToolkit.MVVM.RelayCommand]::new(
            [System.Action[object]]{

                if($null -ne $vm.CloseAction){

                    & $vm.CloseAction $false

                }

            }.GetNewClosure()
        )

    }



    [hashtable] GetValues(){

        $result =
            @{}



        foreach($field in $this.Fields){

            $result[$field.Label.Value] =
                $field.Value.Value

        }


        return $result

    }

}