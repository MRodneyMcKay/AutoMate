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

using namespace System.Windows.Controls


Register-Behavior "SelectedItemsSync" {

    param($element)

    if ($element -isnot [ListBox]) {
        throw "SelectedItemsSync requires ListBox"
    }


    $list = [ListBox]$element


    $list.Add_SelectionChanged({

        param($sender,$args)


        $callback = $sender.Tag

        if($callback -and $callback -is [scriptblock])
        {
            & $callback @($sender.SelectedItems)
        }

    }.GetNewClosure())

}