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
using namespace System.Windows.Media


Register-Behavior "ClickEmptyListSpaceCancel" {

    param($element)


    if($element -isnot [ListBox])
    {
        throw "ClickEmptyListSpaceCancel requires ListBox"
    }


    $listBox = [ListBox]$element


    $listBox.Add_PreviewMouseDown({

        param($sender,$args)


        $source = $args.OriginalSource


        while($source)
        {

            if($source -is [ListBoxItem])
            {
                return
            }


            if($source -eq $sender)
            {
                break
            }


            $source =
                [VisualTreeHelper]::GetParent($source)
        }


        $action =
            [WPFToolkit.Behaviors.BehaviorBridge]::GetAction($sender)


        if($action -is [scriptblock])
        {
            & $action
        }


    }.GetNewClosure())

}