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


    if ($element -isnot [System.Windows.Controls.ListBox]) {
        throw "DragSelection requires ListBox, got $($element.GetType().FullName)"
    }

    $listBox = [System.Windows.Controls.ListBox]$element


    $getItemAtPoint = {
        param(
            [System.Windows.Controls.ListBox]$ListBox,
            [System.Windows.Point]$Point
        )

        $visual = $ListBox.InputHitTest($Point)

        while(
            $visual -and
            -not ($visual -is [System.Windows.Controls.ListBoxItem])
        )
        {
            $visual = [System.Windows.Media.VisualTreeHelper]::GetParent($visual)
        }

        return $visual
    }.GetNewClosure()


    $state = @{
        Dragging = $false
        Anchor   = -1
    }

    $ctrlClickState = @{
        IsCtrlClick       = $false
        ClickedItemWasSelected = $false
        PreviousSelection = @()
    }


    $listBox.AddHandler(
        [System.Windows.UIElement]::PreviewMouseLeftButtonDownEvent,
        [System.Windows.Input.MouseButtonEventHandler]{

            param($sender,$mouseEvent)

            if(
                [System.Windows.Input.Keyboard]::Modifiers -ne [System.Windows.Input.ModifierKeys]::None
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
        [System.Windows.UIElement]::PreviewMouseLeftButtonDownEvent,
        [System.Windows.Input.MouseButtonEventHandler]{
            param($sender, $mouseEvent)
            $ctrlClickState.IsCtrlClick = (
                ([System.Windows.Input.Keyboard]::Modifiers -band
                    [System.Windows.Input.ModifierKeys]::Control) -ne 0
            )
            $ctrlClickState.ClickedItemWasSelected = $false
            $ctrlClickState.PreviousSelection = @()

            if (-not $ctrlClickState.IsCtrlClick) {
                return
            }

            $item = & $getItemAtPoint $sender $mouseEvent.GetPosition($sender)
            if ($null -eq $item) {
                $ctrlClickState.IsCtrlClick = $false
                return
            }

            $clickedContent = $item.Content
            $ctrlClickState.ClickedItemWasSelected =
                $sender.SelectedItems.Contains($clickedContent)
            $ctrlClickState.PreviousSelection = @($sender.SelectedItems)
        }.GetNewClosure(),
        $true
    )

    $listBox.AddHandler(
        [System.Windows.UIElement]::PreviewMouseLeftButtonUpEvent,
        [System.Windows.Input.MouseButtonEventHandler]{
            param($sender, $mouseEvent)
            if (-not $ctrlClickState.IsCtrlClick -or
                $ctrlClickState.ClickedItemWasSelected) {
                $ctrlClickState.IsCtrlClick = $false
                return
            }

            foreach ($item in $ctrlClickState.PreviousSelection) {
                if (-not $sender.SelectedItems.Contains($item)) {
                    [void]$sender.SelectedItems.Add($item)
                }
            }
            $ctrlClickState.IsCtrlClick = $false
        }.GetNewClosure(),
        $true
    )



    $listBox.AddHandler(
        [System.Windows.UIElement]::PreviewMouseMoveEvent,
        [System.Windows.Input.MouseEventHandler]{

            param($sender,$mouseEvent)


            if(-not $state.Dragging)
            {
                return
            }


            if(
                $mouseEvent.LeftButton -ne
                [System.Windows.Input.MouseButtonState]::Pressed
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
        [System.Windows.UIElement]::PreviewMouseLeftButtonUpEvent,
        [System.Windows.Input.MouseButtonEventHandler]{

            param($sender,$mouseEvent)

            $state.Dragging = $false

        }.GetNewClosure(),
        $true
    )

}
