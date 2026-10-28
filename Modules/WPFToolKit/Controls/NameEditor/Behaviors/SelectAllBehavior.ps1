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
using namespace System.Windows.Controls
using namespace System.Windows.Input


Register-Behavior "SelectAll" {

    param($element)

    if ($element -isnot [System.Windows.Controls.ListBox]) {
        throw "SelectAll requires ListBox, got $($element.GetType().FullName)"
    }

    $listBox = [System.Windows.Controls.ListBox]$element


    $listBox.AddHandler(
        [System.Windows.UIElement]::PreviewKeyDownEvent,
        [System.Windows.Input.KeyEventHandler]{

            param($sender,$args)


            if(
                $args.Key -eq                 [System.Windows.Input.Key]::A -and
                [System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control
            )
            {
                $sender.SelectAll()
                $args.Handled = $true
            }

        }.GetNewClosure(),
        $true
    )

}