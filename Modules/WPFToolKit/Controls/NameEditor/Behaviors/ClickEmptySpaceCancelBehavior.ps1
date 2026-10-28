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


    if($element -isnot [System.Windows.Controls.ListBox])
    {
        throw "ClickEmptyListSpaceCancel requires ListBox"
    }


    $listBox = [System.Windows.Controls.ListBox]$element


    $listBox.Add_PreviewMouseDown({

        param($sender,$args)


        $source = $args.OriginalSource


        while($source)
        {

            if($source -is [System.Windows.Controls.ListBoxItem])
            {
                return
            }


            if($source -eq $sender)
            {
                break
            }


            $source =
                [System.Windows.Media.VisualTreeHelper]::GetParent($source)
        }


        $action =
            [WPFToolkit.Behaviors.BehaviorBridge]::GetAction($sender)


        if($action -is [scriptblock])
        {
            & $action
        }


    }.GetNewClosure())

}