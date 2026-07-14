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
using namespace System.Windows.Markup

function Get-ListBoxItemAtPoint {
    param (
        [System.Windows.Controls.ListBox]$ListBox,
        [System.Windows.Point]$Point
    )

    $Element = $ListBox.InputHitTest($Point)
    while ($Element -and -not ($Element -is [System.Windows.Controls.ListBoxItem])) {
        $Element = [System.Windows.Media.VisualTreeHelper]::GetParent($Element)
    }
    return $Element
}

function Enable-ListBoxDragSelection {
    param (
        [System.Windows.Controls.ListBox]$ListBox
    )

    $DragState = @{
        IsDragging  = $false
        AnchorIndex = -1
        LastStart   = -1
        LastEnd     = -1
    }

    $ListBox.Add_PreviewMouseLeftButtonDown({
        param ($Sender, $Event)

        if ([System.Windows.Input.Keyboard]::Modifiers -ne [System.Windows.Input.ModifierKeys]::None) { return }

        $Item = Get-ListBoxItemAtPoint -ListBox $Sender -Point $Event.GetPosition($Sender)
        if ($null -eq $Item) { return }

        $Index = $Sender.ItemContainerGenerator.IndexFromContainer($Item)
        if ($Index -lt 0) { return }

        $DragState.IsDragging  = $true
        $DragState.AnchorIndex = $Index
        $DragState.LastStart   = $Index
        $DragState.LastEnd     = $Index
    }.GetNewClosure())

    $ListBox.Add_PreviewMouseMove({
        param ($Sender, $Event)

        if (-not $DragState.IsDragging) { return }
        if ($Event.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) {
            $DragState.IsDragging = $false
            return
        }

        $Item = Get-ListBoxItemAtPoint -ListBox $Sender -Point $Event.GetPosition($Sender)
        if ($null -eq $Item) { return }

        $Index = $Sender.ItemContainerGenerator.IndexFromContainer($Item)
        if ($Index -lt 0) { return }

        $Start = [Math]::Min($DragState.AnchorIndex, $Index)
        $End   = [Math]::Max($DragState.AnchorIndex, $Index)

        if ($Start -eq $DragState.LastStart -and $End -eq $DragState.LastEnd) { return }

        $DragState.LastStart = $Start
        $DragState.LastEnd   = $End

        $Sender.SelectedItems.Clear()
        for ($i = $Start; $i -le $End; $i++) {
            [void]$Sender.SelectedItems.Add($Sender.Items[$i])
        }
    }.GetNewClosure())

    $ListBox.Add_PreviewMouseLeftButtonUp({
        param ($Sender, $Event)
        $DragState.IsDragging = $false
    }.GetNewClosure())
}


function Get-ListBoxItemAtPoint {
    param (
        [ListBox]$ListBox,
        [Point]$Point
    )

    $Element = $ListBox.InputHitTest($Point)
    while ($Element -and -not ($Element -is [ListBoxItem])) {
        $Element = [VisualTreeHelper]::GetParent($Element)
    }
    return $Element
}

function Enable-ListBoxDragSelection {
    param (
        [ListBox]$ListBox
    )

    $DragState = @{
        IsDragging  = $false
        AnchorIndex = -1
        LastStart   = -1
        LastEnd     = -1
    }

    $ListBox.Add_PreviewMouseLeftButtonDown({
        param ($Sender, $Event)

        if ([Keyboard]::Modifiers -ne [ModifierKeys]::None) { return }

        $Item = Get-ListBoxItemAtPoint -ListBox $Sender -Point $Event.GetPosition($Sender)
        if ($null -eq $Item) { return }

        $Index = $Sender.ItemContainerGenerator.IndexFromContainer($Item)
        if ($Index -lt 0) { return }

        $DragState.IsDragging  = $true
        $DragState.AnchorIndex = $Index
        $DragState.LastStart   = $Index
        $DragState.LastEnd     = $Index
    }.GetNewClosure())

    $ListBox.Add_PreviewMouseMove({
        param ($Sender, $Event)

        if (-not $DragState.IsDragging) { return }
        if ($Event.LeftButton -ne [MouseButtonState]::Pressed) {
            $DragState.IsDragging = $false
            return
        }

        $Item = Get-ListBoxItemAtPoint -ListBox $Sender -Point $Event.GetPosition($Sender)
        if ($null -eq $Item) { return }

        $Index = $Sender.ItemContainerGenerator.IndexFromContainer($Item)
        if ($Index -lt 0) { return }

        $Start = [Math]::Min($DragState.AnchorIndex, $Index)
        $End   = [Math]::Max($DragState.AnchorIndex, $Index)

        if ($Start -eq $DragState.LastStart -and $End -eq $DragState.LastEnd) { return }

        $DragState.LastStart = $Start
        $DragState.LastEnd   = $End

        $Sender.SelectedItems.Clear()
        for ($i = $Start; $i -le $End; $i++) {
            [void]$Sender.SelectedItems.Add($Sender.Items[$i])
        }
    }.GetNewClosure())

    $ListBox.Add_PreviewMouseLeftButtonUp({
        param ($Sender, $Event)
        $DragState.IsDragging = $false
    }.GetNewClosure())
}

function New-NameEditorControl {
    param (
        [string]$XamlPath,
        [Window]$HostWindow = $null
    )

    if (-not (Test-Path -Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }

    $xaml = Get-Content -Path $XamlPath -Raw
    $root = [XamlReader]::Parse($xaml)

    $nameList       = $root.FindName('NameList')
    $inputBox       = $root.FindName('InputBox')
    $cancelButton   = $root.FindName('CancelButton')
    $addUpdateButton= $root.FindName('AddUpdateButton')
    $deleteButton   = $root.FindName('DeleteButton')

    Enable-ListBoxDragSelection -ListBox $nameList

    $editor = [PSCustomObject]@{
        Control            = $root
        NameList           = $nameList
        InputBox           = $inputBox
        CancelButton       = $cancelButton
        AddUpdateButton    = $addUpdateButton
        DeleteButton       = $deleteButton

        OnAdd              = $null
        OnUpdate           = $null
        OnDelete           = $null
        OnCancel           = $null
        OnSelectionChanged = $null
    }

    $inputBox.Add_TextChanged({
        param($sender, $event)

        $trimmed = $sender.Text.TrimEnd("`r", "`n")
        if ($trimmed -ne $sender.Text) {
            $caret = $trimmed.Length
            $sender.Text = $trimmed
            $sender.CaretIndex = $caret
        }

        $cancelButton.Visibility = 
            if ([string]::IsNullOrWhiteSpace($sender.Text)) { 'Collapsed' } else { 'Visible' }
    }.GetNewClosure())

    $nameList.Add_SelectionChanged({
        param ($Sender, $Event)

        $selected = @($nameList.SelectedItems | ForEach-Object { $_ })

        if ($editor.OnSelectionChanged) {
            & $editor.OnSelectionChanged $selected
        }

        switch ($nameList.SelectedItems.Count) {
            0 {
                $inputBox.Clear()
                $inputBox.Visibility    = 'Visible'
                $addUpdateButton.Visibility = 'Visible'
                $cancelButton.Visibility    = 'Collapsed'
                $deleteButton.Visibility    = 'Collapsed'
                $addUpdateButton.Content    = 'Hozzáad'
            }
            1 {
                $inputBox.Visibility    = 'Visible'
                $addUpdateButton.Visibility = 'Visible'
                $inputBox.Text          = $nameList.SelectedItem
                $addUpdateButton.Content= 'Frissítés'
                $deleteButton.Visibility= 'Visible'
            }
            default {
                $inputBox.Clear()
                $inputBox.Visibility    = 'Collapsed'
                $addUpdateButton.Visibility = 'Collapsed'
                $cancelButton.Visibility    = 'Collapsed'
                $deleteButton.Visibility    = 'Visible'
            }
        }
    }.GetNewClosure())

    $addUpdateButton.Add_Click({
        $text = $inputBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return }

        if ($null -ne $nameList.SelectedItem -and $nameList.SelectedItems.Count -eq 1) {
            if ($editor.OnUpdate) {
                & $editor.OnUpdate $nameList.SelectedItem $text
            }
        }
        else {
            if ($editor.OnAdd) {
                & $editor.OnAdd $text
            }
        }
    }.GetNewClosure())

    $deleteButton.Add_Click({
        if ($nameList.SelectedItems.Count -eq 0) { return }
        $names = @($nameList.SelectedItems | ForEach-Object { [string]$_ })
        if ($editor.OnDelete) {
            & $editor.OnDelete $names
        }
    }.GetNewClosure())

    $cancelButton.Add_Click({
        $inputBox.Clear()
        $nameList.UnselectAll()
        if ($editor.OnCancel) {
            & $editor.OnCancel
        }
    }.GetNewClosure())

    $inputBox.Add_KeyDown({
        param ($Sender, $Event)
        if ($Event.Key -eq 'Enter') {
            $text = $inputBox.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($text)) {
                $Event.Handled = $true
                return
            }

            if ($null -ne $nameList.SelectedItem -and $nameList.SelectedItems.Count -eq 1) {
                if ($editor.OnUpdate) {
                    & $editor.OnUpdate $nameList.SelectedItem $text
                }
            }
            else {
                if ($editor.OnAdd) {
                    & $editor.OnAdd $text
                }
            }
            $Event.Handled = $true
        }
    }.GetNewClosure())

    if ($HostWindow) {
        $HostWindow.Add_PreviewMouseDown({
            param($Sender, $Event)

            $source = $Event.OriginalSource
            while ($source) {
                if ($source -is [ListBoxItem]) { return }
                if ($source -is [Button] -or
                    $source -is [TextBox] -or
                    $source -is [TabItem] -or
                    $source -is [MenuItem] -or
                    $source -is [ContextMenu]) {
                    return
                }
                $source = [VisualTreeHelper]::GetParent($source)
            }

            $inputBox.Clear()
            $nameList.UnselectAll()
            if ($editor.OnCancel) {
                & $editor.OnCancel
            }
        }.GetNewClosure())

        $HostWindow.Add_PreviewKeyDown({
            param($Sender, $Event)
            if ($Event.Key -eq [Key]::Escape) {
                $inputBox.Clear()
                $nameList.UnselectAll()
                if ($editor.OnCancel) {
                    & $editor.OnCancel
                }
                $Event.Handled = $true
            }
        }.GetNewClosure())
    }

    return $editor
}

function Get-ListBoxItemAtPoint {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Controls.ListBox] $ListBox,

        [Parameter(Mandatory)]
        [System.Windows.Point] $Point
    )

    # Hit test
    $hit = [System.Windows.Media.VisualTreeHelper]::HitTest($ListBox, $Point)
    if ($hit -eq $null) { return $null }

    # Walk up the visual tree until we find a ListBoxItem
    $current = $hit.VisualHit
    while ($current -ne $null -and -not ($current -is [System.Windows.Controls.ListBoxItem])) {
        $current = [System.Windows.Media.VisualTreeHelper]::GetParent($current)
    }

    return $current
}
