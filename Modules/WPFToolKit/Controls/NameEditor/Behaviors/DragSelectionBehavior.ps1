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
using namespace System.Windows.Media


Register-Behavior "DragSelection" {

    param($element)


    if ($element -isnot [ListBox]) {
        throw "DragSelection requires ListBox, got $($element.GetType().FullName)"
    }

    $listBox = [ListBox]$element


    $getItemAtPoint = {
        param(
            [ListBox]$ListBox,
            [System.Windows.Point]$Point
        )

        $visual = $ListBox.InputHitTest($Point)

        while(
            $visual -and
            -not ($visual -is [ListBoxItem])
        )
        {
            $visual = [VisualTreeHelper]::GetParent($visual)
        }

        return $visual
    }.GetNewClosure()


    $state = @{
        Dragging = $false
        Anchor   = -1
    }


    $listBox.AddHandler(
        [UIElement]::PreviewMouseLeftButtonDownEvent,
        [MouseButtonEventHandler]{

            param($sender,$mouseEvent)

            if(
                [Keyboard]::Modifiers -ne [ModifierKeys]::None
            )
            {
                return
            }


            $item = & $getItemAtPoint `
                $sender `
                $mouseEvent.GetPosition($sender)


            if($null -eq $item)
            {
                return
            }


            $state.Anchor =
                $sender.ItemContainerGenerator.IndexFromContainer($item)

            $state.Dragging = $true

        }.GetNewClosure(),
        $true
    )



    $listBox.AddHandler(
        [UIElement]::PreviewMouseMoveEvent,
        [MouseEventHandler]{

            param($sender,$mouseEvent)


            if(-not $state.Dragging)
            {
                return
            }


            if(
                $mouseEvent.LeftButton -ne
                [MouseButtonState]::Pressed
            )
            {
                $state.Dragging = $false
                return
            }


            $item = & $getItemAtPoint `
                $sender `
                $mouseEvent.GetPosition($sender)


            if($null -eq $item)
            {
                return
            }


            $index =
                $sender.ItemContainerGenerator.IndexFromContainer($item)


            if($index -lt 0)
            {
                return
            }


            $start = [Math]::Min(
                $state.Anchor,
                $index
            )

            $end = [Math]::Max(
                $state.Anchor,
                $index
            )


            for($i=$start; $i -le $end; $i++)
            {
                $container =
                    $sender.ItemContainerGenerator.ContainerFromIndex($i)

                if($container)
                {
                    $container.IsSelected = $true
                }
            }

        }.GetNewClosure(),
        $true
    )



    $listBox.AddHandler(
        [UIElement]::PreviewMouseLeftButtonUpEvent,
        [MouseButtonEventHandler]{

            param($sender,$mouseEvent)

            $state.Dragging = $false

        }.GetNewClosure(),
        $true
    )

}
