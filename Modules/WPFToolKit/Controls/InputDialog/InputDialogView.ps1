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


class InputDialogView {


    [Window]
    $Window


    [InputDialogViewModel]
    $ViewModel



    InputDialogView(
        [InputDialogViewModel] $viewModel
    ){

        $this.ViewModel =
            $viewModel



        $xaml =
            Get-Content `
                (Join-Path $PSScriptRoot "InputDialogView.xaml") `
                -Raw



        $this.Window =
            [Window][XamlReader]::Parse(
                $xaml
            )



        $this.Window.DataContext =
            $this.ViewModel



        $dialog =
            $this.Window



        $this.ViewModel.CloseAction = {

            param(
                [bool] $result
            )


            $dialog.DialogResult =
                $result


        }.GetNewClosure()

    }



    [bool] Show(){

        return [bool]$this.Window.ShowDialog()

    }

}