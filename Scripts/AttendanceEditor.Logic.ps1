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

function Sort-Names {
    param (
        [System.Collections.ObjectModel.ObservableCollection[string]]$Collection
    )

    $Sorted = $Collection | Sort-Object -Culture $global:CultureHU
    $Collection.Clear()
    foreach ($Name in $Sorted) {
        $Collection.Add($Name)
    }
}

function Add-TabItemBeforePlusTab {
    param (
        [System.Windows.Controls.TabControl]$TabControl,
        [System.Windows.Controls.TabItem]$NewTab
    )

    $PlusTab = $TabControl.Tag
    if ($PlusTab -and $TabControl.Items.Contains($PlusTab)) {
        $Index = $TabControl.Items.IndexOf($PlusTab)
        [void]($TabControl.Items.Insert($Index, $NewTab))
    }
    else {
        [void]($TabControl.Items.Add($NewTab))
    }
}

function Mark-Dirty {
    param (
        [string]$DeptName,
        [string]$PosName = $null
    )

    if ($global:DirtyDepartments.Add($DeptName)) {
        if ($global:DeptControls.ContainsKey($DeptName)) {
            $Facility = $global:Data[$DeptName].Facility
            Set-HeaderDirtyState -HeaderBlock $global:DeptControls[$DeptName].HeaderBlock -Name $DeptName -Facility $Facility -IsDirty $true
        }
    }

    if ($PosName) {
        $Key = "$DeptName|$PosName"
        if ($global:DirtyPositions.Add($Key)) {
            if ($global:DeptControls.ContainsKey($DeptName) -and $global:DeptControls[$DeptName].PositionControls.ContainsKey($PosName)) {
                Set-HeaderDirtyState -HeaderBlock $global:DeptControls[$DeptName].PositionControls[$PosName].HeaderBlock -Name $PosName -IsDirty $true
            }
        }
    }

    Update-ReloadButtonVisibility
}

function Update-ReloadButtonVisibility {
    if ($global:DirtyDepartments.Count -gt 0 -or $global:DirtyPositions.Count -gt 0) {
        $global:ReloadButton.Visibility = "Visible"
    }
    else {
        $global:ReloadButton.Visibility = "Collapsed"
    }
}

function Clear-AllDirty {
    foreach ($DeptName in $global:DirtyDepartments) {
        if ($global:DeptControls.ContainsKey($DeptName)) {
            $Facility = $global:Data[$DeptName].Facility
            Set-HeaderDirtyState -HeaderBlock $global:DeptControls[$DeptName].HeaderBlock -Name $DeptName -Facility $Facility -IsDirty $false
        }
    }
    $global:DirtyDepartments.Clear()

    foreach ($Key in $global:DirtyPositions) {
        $Parts = $Key -split '\|', 2
        $DeptName = $Parts[0]
        $PosName = $Parts[1]
        if ($global:DeptControls.ContainsKey($DeptName) -and $global:DeptControls[$DeptName].PositionControls.ContainsKey($PosName)) {
            Set-HeaderDirtyState -HeaderBlock $global:DeptControls[$DeptName].PositionControls[$PosName].HeaderBlock -Name $PosName -IsDirty $false
        }
    }
    $global:DirtyPositions.Clear()

    Update-ReloadButtonVisibility
}

function Get-CurrentDeptAndPosition {
    $CurrentDeptTab = $global:TabControl.SelectedItem
    if ($null -eq $CurrentDeptTab) { return $null }

    $DeptName = $CurrentDeptTab.Tag
    if (-not $global:DeptControls.ContainsKey($DeptName)) { return $null }

    $PositionTabControl = $global:DeptControls[$DeptName].PositionTabControl
    $CurrentPosTab = $PositionTabControl.SelectedItem
    if ($null -eq $CurrentPosTab) { return $null }

    $PosName = $CurrentPosTab.Tag
    if (-not $global:DeptControls[$DeptName].PositionControls.ContainsKey($PosName)) { return $null }

    return [PSCustomObject]@{ DeptName = $DeptName; PosName = $PosName }
}

function Add-NameToCurrentPosition {
    $Context = Get-CurrentDeptAndPosition
    if ($null -eq $Context) {
        $global:Status.Text = "Nincs kiválasztott pozíció"
        return
    }

    $DeptName = $Context.DeptName
    $PosName  = $Context.PosName
    $Controls = $global:DeptControls[$DeptName].PositionControls[$PosName]

    $Collection = $global:Data[$DeptName].Positions[$PosName]

    # Split on tabs and line breaks (Excel clipboard format)
    $Names = $Controls.Box.Text `
        -split "(`r`n|`n|`r|`t)" |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique

    if ($Names.Count -eq 0) {
        return
    }

    $Added = 0
    $Duplicates = 0

    foreach ($NewName in $Names) {

        if ($Collection -contains $NewName) {
            $Duplicates++
            Write-Log -Message "Duplikált név: $DeptName / $PosName / $NewName" -Level "WARNING"
            continue
        }

        $Collection.Add($NewName)
        $Added++

        Write-Log -Message "Név hozzáadva: $DeptName / $PosName / $NewName" -Level "INFO"
    }

    if ($Added -gt 0) {
        Sort-Names -Collection $Collection
        Mark-Dirty -DeptName $DeptName -PosName $PosName
    }

    $Controls.Box.Clear()

    if ($Added -gt 0 -and $Duplicates -eq 0) {
        $global:Status.Text = "$Added név hozzáadva"
    }
    elseif ($Added -gt 0) {
        $global:Status.Text = "$Added hozzáadva, $Duplicates duplikált"
    }
    else {
        $global:Status.Text = "Minden név már létezik"
    }
}

function Remove-NameFromCurrentPosition {
    $Context = Get-CurrentDeptAndPosition
    if ($null -eq $Context) {
        $global:Status.Text = "Nincs kiválasztott pozíció"
        return
    }

    $DeptName = $Context.DeptName
    $PosName  = $Context.PosName
    $Controls = $global:DeptControls[$DeptName].PositionControls[$PosName]

    if ($Controls.List.SelectedItems.Count -eq 0) {
        $global:Status.Text = "Nincs kiválasztott név"
        return
    }

    # Copy because SelectedItems changes while removing
    $Names = @($Controls.List.SelectedItems)

    foreach ($Name in $Names) {
        $global:Data[$DeptName].Positions[$PosName].Remove($Name)
        Write-Log -Message "Törölve: $DeptName / $PosName / $Name" -Level "INFO"
    }

    Mark-Dirty -DeptName $DeptName -PosName $PosName

    $Controls.Box.Clear()
    $Controls.List.UnselectAll()

    $global:Status.Text = "$($Names.Count) név törölve"
}

function Update-NameInCurrentPosition {
    $Context = Get-CurrentDeptAndPosition
    if ($null -eq $Context) {
        $global:Status.Text = "Nincs kiválasztott pozíció"
        return
    }

    $DeptName = $Context.DeptName
    $PosName  = $Context.PosName
    $Controls = $global:DeptControls[$DeptName].PositionControls[$PosName]

    $OldName = $Controls.List.SelectedItem
    if ($null -eq $OldName) {
        $global:Status.Text = "Nincs kiválasztott név"
        return
    }

    $NewName = $Controls.Box.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($NewName)) {
        $global:Status.Text = "A név nem lehet üres"
        return
    }

    $Collection = $global:Data[$DeptName].Positions[$PosName]

    if ($NewName -eq $OldName) {
        $Controls.Box.Clear()
        $Controls.List.SelectedItem = $null
        $global:Status.Text = "Nincs változás"
        return
    }

    if ($Collection -contains $NewName) {
        Write-Log -Message "Frissítés sikertelen, duplikált név: $DeptName / $PosName / $NewName" -Level "WARNING"
        $global:Status.Text = "Ez a név már létezik"
        return
    }

    $Index = $Collection.IndexOf($OldName)
    if ($Index -lt 0) {
        $global:Status.Text = "A név már nem található"
        return
    }

    $Collection[$Index] = $NewName
    Sort-Names -Collection $Collection
    Mark-Dirty -DeptName $DeptName -PosName $PosName

    Write-Log -Message "Név módosítva: $DeptName / $PosName / $OldName -> $NewName" -Level "INFO"
    $global:Status.Text = "Név módosítva: $NewName"

    $Controls.Box.Clear()
    $Controls.List.SelectedItem = $null
}

function Cancel-NameEdit {
    $Context = Get-CurrentDeptAndPosition
    if ($null -eq $Context) {
        return
    }

    $Controls = $global:DeptControls[$Context.DeptName].PositionControls[$Context.PosName]

    $Controls.Box.Clear()
    $Controls.List.SelectedItem = $null
}

function Confirm-Deletion {
    param (
        [string]$Message,
        [string]$Title = "Megerősítés"
    )

    $result = [System.Windows.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    return $result -eq [System.Windows.MessageBoxResult]::Yes
}

function Delete-Position {
    param (
        [string]$DeptName,
        [string]$PosName
    )

    if (-not $global:Data.Contains($DeptName)) {
        Write-Log -Message "Pozíció törlése sikertelen, ismeretlen részleg: $DeptName" -Level "WARNING"
        return
    }

    if (-not $global:Data[$DeptName].Positions.Contains($PosName)) {
        Write-Log -Message "Pozíció törlése sikertelen, ismeretlen pozíció: $DeptName / $PosName" -Level "WARNING"
        return
    }

    if (-not (Confirm-Deletion -Message "Biztosan törlöd a pozíciót: $PosName?" -Title "Pozíció törlése")) {
        return
    }

    $global:Data[$DeptName].Positions.Remove($PosName)

    if ($global:DeptControls.ContainsKey($DeptName) -and $global:DeptControls[$DeptName].PositionControls.ContainsKey($PosName)) {
        $Tab = $global:DeptControls[$DeptName].PositionControls[$PosName].Tab
        $global:DeptControls[$DeptName].PositionControls.Remove($PosName)
        if ($Tab -and $global:DeptControls[$DeptName].PositionTabControl.Items.Contains($Tab)) {
            $global:DeptControls[$DeptName].PositionTabControl.Items.Remove($Tab)
        }
    }
    else {
        Write-Log -Message "Pozíció törölve az adatból, de UI vezérlő nem található: $DeptName / $PosName" -Level "WARNING"
    }

    Write-Log -Message "Pozíció törölve: $DeptName / $PosName" -Level "INFO"
    $global:Status.Text = "Pozíció törölve: $PosName"
    Save-All
}

function Delete-Department {
    param (
        [string]$DeptName
    )

    if (-not $global:Data.Contains($DeptName)) {
        Write-Log -Message "Részleg törlése sikertelen, ismeretlen részleg: $DeptName" -Level "WARNING"
        return
    }

    if (-not (Confirm-Deletion -Message "Biztosan törlöd a részleget: $DeptName és minden pozícióját?" -Title "Részleg törlése")) {
        return
    }

    if ($global:DeptControls.ContainsKey($DeptName)) {
        $Tab = $global:DeptControls[$DeptName].Tab
        $global:DeptControls.Remove($DeptName)
        if ($Tab -and $global:TabControl.Items.Contains($Tab)) {
            $global:TabControl.Items.Remove($Tab)
        }
    }
    else {
        Write-Log -Message "Részleg törölve az adatból, de UI vezérlő nem található: $DeptName" -Level "WARNING"
    }

    $global:Data.Remove($DeptName)

    Write-Log -Message "Részleg törölve: $DeptName" -Level "INFO"
    $global:Status.Text = "Részleg törölve: $DeptName"
    Save-All
}

function New-PositionTab {
    param (
        [string]$DeptName,
        [string]$PosName
    )

    $PositionTabControl = $global:DeptControls[$DeptName].PositionTabControl

    $Tab = New-Object System.Windows.Controls.TabItem
    $Tab.Tag = $PosName
    $HeaderBlock = New-TabHeaderBlock -Name $PosName
    $Tab.Header = $HeaderBlock

    $ContextMenu = New-Object System.Windows.Controls.ContextMenu
    $DeleteMenuItem = New-Object System.Windows.Controls.MenuItem
    $DeleteMenuItem.Header = "Törlés"
    
    # Capture values in local variables to avoid closure issues
    $CapturedDeptName = $DeptName
    $CapturedPosName = $PosName
    
    $DeleteMenuItem.Add_Click({
        param ($Sender, $Event)
        try {
            Delete-Position -DeptName $CapturedDeptName -PosName $CapturedPosName
        }
        catch {
            Write-Log -Message "Hiba a pozició törlésekor: $($_.Exception.Message)" -Level "ERROR"
            $global:Status.Text = "Törlési hiba"
        }
    }.GetNewClosure())
    [void]($ContextMenu.Items.Add($DeleteMenuItem))
    $HeaderBlock.ContextMenu = $ContextMenu

    $Grid = New-Object System.Windows.Controls.Grid
    $Grid.Margin = "0,10,0,0"
    
    $ListRow = New-Object Windows.Controls.RowDefinition
    $ListRow.Height = "*"
    $InputRow = New-Object Windows.Controls.RowDefinition
    $InputRow.Height = "Auto"

    $Grid.RowDefinitions.Add($ListRow)
    $Grid.RowDefinitions.Add($InputRow)

    $List = New-Object System.Windows.Controls.ListBox
    $List.SelectionMode = "Extended"
    $List.Margin = "0,0,0,15"
    $List.ItemsSource = $global:Data[$DeptName].Positions[$PosName]
    [Windows.Controls.Grid]::SetRow($List, 0)

    $Panel = New-Object System.Windows.Controls.StackPanel
    $Panel.Orientation = "Horizontal"

    $Box = New-Object System.Windows.Controls.TextBox
    $Box.Width = 300
    $Box.AcceptsReturn = $true
    $Box.TextWrapping = "Wrap"
    $Box.VerticalScrollBarVisibility = "Auto"
    $Box.VerticalContentAlignment = "Top"

    $MegseButton = New-Object System.Windows.Controls.Button
    $MegseButton.Content = "Mégse"
    $MegseButton.Width = 100
    $MegseButton.Margin = "10,0,0,0"
    $MegseButton.Visibility = "Collapsed"

    $AddButton = New-Object System.Windows.Controls.Button
    $AddButton.Content = "Hozzáad"
    $AddButton.Width = 100
    $AddButton.Margin = "10,0,0,0"
    $AddButton.Style = $global:Window.FindResource("PrimaryButton")

    $DeleteButton = New-Object System.Windows.Controls.Button
    $DeleteButton.Content = "Törlés"
    $DeleteButton.Height = 30
    $DeleteButton.Width = 100
    $DeleteButton.Margin = "10,0,0,0"
    $DeleteButton.Visibility = "Collapsed"

    $Box.Add_TextChanged({
        param($sender, $event)

        # Remove trailing CR/LF characters
        $trimmed = $sender.Text.TrimEnd("`r", "`n")

        if ($trimmed -ne $sender.Text) {
            $caret = $trimmed.Length
            $sender.Text = $trimmed
            $sender.CaretIndex = $caret
        }

        if ([string]::IsNullOrWhiteSpace($sender.Text)) {
            $MegseButton.Visibility = "Collapsed"
        }
        else {
            $MegseButton.Visibility = "Visible"
        }
    }.GetNewClosure())

    $List.Add_SelectionChanged({
        param ($Sender, $Event)

        switch ($List.SelectedItems.Count) {

            0 {
                $Box.Clear()

                $Box.Visibility = "Visible"
                $AddButton.Visibility = "Visible"
                $MegseButton.Visibility = "Collapsed"
                $DeleteButton.Visibility = "Collapsed"

                $AddButton.Content = "Hozzáad"
            }

            1 {
                $Box.Visibility = "Visible"
                $AddButton.Visibility = "Visible"

                $Box.Text = $List.SelectedItem
                $AddButton.Content = "Frissítés"
                $DeleteButton.Visibility = "Visible"
            }

            default {
                $Box.Clear()

                $Box.Visibility = "Collapsed"
                $AddButton.Visibility = "Collapsed"
                $MegseButton.Visibility = "Collapsed"

                $DeleteButton.Visibility = "Visible"
            }
        }
    }.GetNewClosure())

    $AddButton.Add_Click({
        if ($null -ne $List.SelectedItem) {
            Update-NameInCurrentPosition
        }
        else {
            Add-NameToCurrentPosition
        }
    }.GetNewClosure())

    $DeleteButton.Add_Click({ Remove-NameFromCurrentPosition })

    $MegseButton.Add_Click({ Cancel-NameEdit })

    $Box.Add_KeyDown({
        param ($Sender, $Event)
        if ($Event.Key -eq "Enter") {
            if ($null -ne $List.SelectedItem) {
                Update-NameInCurrentPosition
            }
            else {
                Add-NameToCurrentPosition
            }
            $Event.Handled = $true
        }
    }.GetNewClosure())

    [void]($Panel.Children.Add($Box))
    [void]($Panel.Children.Add($MegseButton))
    [void]($Panel.Children.Add($AddButton))
    [void]($Panel.Children.Add($DeleteButton))
    [Windows.Controls.Grid]::SetRow($Panel, 1)

    [void]($Grid.Children.Add($List))
    [void]($Grid.Children.Add($Panel))
    $Tab.Content = $Grid

    Add-TabItemBeforePlusTab -TabControl $PositionTabControl -NewTab $Tab

    $global:DeptControls[$DeptName].PositionControls[$PosName] = @{
        Tab          = $Tab
        HeaderBlock  = $HeaderBlock
        Box          = $Box
        AddButton    = $AddButton
        DeleteButton = $DeleteButton
        MegseButton  = $MegseButton
        List         = $List
    }
}

function Add-Position {
    $CurrentDeptTab = $global:TabControl.SelectedItem
    if ($null -eq $CurrentDeptTab) {
        $global:Status.Text = "Nincs kiválasztott részleg"
        return
    }

    $DeptName = $CurrentDeptTab.Tag
    $PosName = Show-ModernInputBox -Title "Új pozíció" -Prompt "Add meg az új pozíció nevét ehhez a részleghez: $DeptName"

    if ([string]::IsNullOrWhiteSpace($PosName)) { return }

    if ($global:Data[$DeptName].Positions.Contains($PosName)) {
        Write-Log -Message "A pozíció már létezik: $DeptName / $PosName" -Level "WARNING"
        $global:Status.Text = "A pozíció már létezik"
        return
    }

    $global:Data[$DeptName].Positions[$PosName] = New-Object System.Collections.ObjectModel.ObservableCollection[string]
    New-PositionTab -DeptName $DeptName -PosName $PosName
    Mark-Dirty -DeptName $DeptName -PosName $PosName

    $global:DeptControls[$DeptName].PositionTabControl.SelectedItem = $global:DeptControls[$DeptName].PositionControls[$PosName].Tab
    Write-Log -Message "Új pozíció hozzáadva: $DeptName / $PosName" -Level "INFO"
    $global:Status.Text = "Pozíció hozzáadva: $PosName"
}

function New-DepartmentTab {
    param (
        [string]$DeptName,
        [string]$Facility = ''
    )

    $Tab = New-Object System.Windows.Controls.TabItem
    $Tab.Tag = $DeptName
    $HeaderBlock = New-TabHeaderBlock -Name $DeptName -Facility $Facility
    $Tab.Header = $HeaderBlock

    $ContextMenu = New-Object System.Windows.Controls.ContextMenu
    $DeleteMenuItem = New-Object System.Windows.Controls.MenuItem
    $DeleteMenuItem.Header = "Törlés"
    
    # Capture values in local variables to avoid closure issues
    $CapturedDeptName = $DeptName
    
    $DeleteMenuItem.Add_Click({
        param ($Sender, $Event)
        try {
            Delete-Department -DeptName $CapturedDeptName
        }
        catch {
            Write-Log -Message "Hiba a részleg törlésekor: $($_.Exception.Message)" -Level "ERROR"
            $global:Status.Text = "Törlési hiba"
        }
    }.GetNewClosure())
    [void]($ContextMenu.Items.Add($DeleteMenuItem))
    $HeaderBlock.ContextMenu = $ContextMenu

    $Container = New-Object System.Windows.Controls.Border
    $Container.Background = $global:Window.FindResource("SurfaceBrush")
    $Container.BorderBrush = $global:Window.FindResource("BorderBrush")
    $Container.BorderThickness = "1"
    $Container.CornerRadius = "8"
    $Container.Padding = "15"
    $Container.Margin = "0,10,0,0"

    $PositionTabControl = New-Object System.Windows.Controls.TabControl

    $AddPositionTab = New-PlusTabItem
    $PositionTabControl.Tag = $AddPositionTab
    [void]($PositionTabControl.Items.Add($AddPositionTab))

    $PositionTabControl.Add_SelectionChanged({
        param ($Sender, $Event)
        if ($Sender.SelectedItem -eq $Sender.Tag) {
            $Previous = if ($Event.RemovedItems.Count -gt 0) { $Event.RemovedItems[0] } else { $null }
            Add-Position
            if ($Sender.SelectedItem -eq $Sender.Tag) {
                $Sender.SelectedItem = $Previous
            }
        }
    })

    $Container.Child = $PositionTabControl
    $Tab.Content = $Container

    Add-TabItemBeforePlusTab -TabControl $global:TabControl -NewTab $Tab

    $global:DeptControls[$DeptName] = @{
        Tab                = $Tab
        HeaderBlock        = $HeaderBlock
        PositionTabControl = $PositionTabControl
        PositionControls   = @{}
    }
}

function Add-Department {
    $Name = Show-ModernInputBox -Title "Új részleg" -Prompt "Add meg az új részleg nevét:"
    if ([string]::IsNullOrWhiteSpace($Name)) { return }

    $Facility = Show-ModernInputBox -Title "Létesítmény" -Prompt "Add meg a részleg létesítményét:"
    if ([string]::IsNullOrWhiteSpace($Facility)) { $Facility = '' }

    if ($global:Data.Contains($Name)) {
        Write-Log -Message "A részleg már létezik: $Name" -Level "WARNING"
        $global:Status.Text = "A részleg már létezik"
        return
    }

    $global:Data[$Name] = [ordered]@{
        Facility = $Facility
        Positions = [ordered]@{}
    }
    New-DepartmentTab -DeptName $Name -Facility $Facility
    Mark-Dirty -DeptName $Name

    $global:TabControl.SelectedItem = $global:DeptControls[$Name].Tab
    Write-Log -Message "Új részleg hozzáadva: $Name" -Level "INFO"
    $global:Status.Text = "Részleg hozzáadva: $Name"
}

function Save-All {
    try {
        $Root = [System.Xml.Linq.XElement]::new([System.Xml.Linq.XName]"Jelenlet")

        foreach ($DeptName in $global:Data.Keys) {
            $DeptElement = [System.Xml.Linq.XElement]::new([System.Xml.Linq.XName]"Department")
            $DeptElement.SetAttributeValue("Name", $DeptName)
            $DeptElement.SetAttributeValue("Facility", $global:Data[$DeptName].Facility)

            foreach ($PosName in $global:Data[$DeptName].Positions.Keys) {
                $PosElement = [System.Xml.Linq.XElement]::new([System.Xml.Linq.XName]"Position")
                $PosElement.SetAttributeValue("Name", $PosName)

                $SortedNames = $global:Data[$DeptName].Positions[$PosName] | Sort-Object -Culture $global:CultureHU

                foreach ($Name in $SortedNames) {
                    $PosElement.Add([System.Xml.Linq.XElement]::new([System.Xml.Linq.XName]"Nev", $Name))
                }
                $DeptElement.Add($PosElement)
            }
            $Root.Add($DeptElement)
        }

        $XDoc = New-Object System.Xml.Linq.XDocument
        $XDoc.Add($Root)
        $XDoc.Save($global:XmlPath)

        Clear-AllDirty

        Write-Log -Message "XML mentve: $global:XmlPath" -Level "INFO"
        $global:Status.Text = "Mentés kész"
    }
    catch {
        Write-Log -Message "Mentési hiba: $($_.Exception.Message)" -Level "ERROR" -ShowMessageBox
        $global:Status.Text = "Mentés sikertelen"
    }
}

function Load-All {
    $global:TabControl.Items.Clear()
    $global:Data.Clear()
    $global:DeptControls.Clear()
    $global:DirtyDepartments.Clear()
    $global:DirtyPositions.Clear()
    Update-ReloadButtonVisibility

    if (-not (Test-Path -Path $global:XmlPath)) {
        Write-Log -Message "Nem található: $global:XmlPath" -Level "ERROR" -ShowMessageBox
        return
    }

    try { $XDoc = [System.Xml.Linq.XDocument]::Load($global:XmlPath) }
    catch {
        Write-Log -Message "Betöltési hiba: $($_.Exception.Message)" -Level "ERROR" -ShowMessageBox
        return
    }

    foreach ($DeptElement in $XDoc.Root.Elements("Department")) {
        $DeptName = $DeptElement.Attribute("Name").Value
        $Facility = ''
        if ($DeptElement.Attribute("Facility")) { $Facility = $DeptElement.Attribute("Facility").Value }
        Write-Log -Message "Betöltés: $DeptName ($Facility)" -Level "INFO"

        $global:Data[$DeptName] = [ordered]@{
            Facility = $Facility
            Positions = [ordered]@{}
        }
        New-DepartmentTab -DeptName $DeptName -Facility $Facility

        foreach ($PosElement in $DeptElement.Elements("Position")) {
            $PosName = $PosElement.Attribute("Name").Value
            $Collection = New-Object System.Collections.ObjectModel.ObservableCollection[string]

            foreach ($NevElement in $PosElement.Elements("Nev")) {
                $NameValue = $NevElement.Value.Trim()
                if ($NameValue) { $Collection.Add($NameValue) }
            }

            Sort-Names -Collection $Collection
            $global:Data[$DeptName].Positions[$PosName] = $Collection
            New-PositionTab -DeptName $DeptName -PosName $PosName
        }
    }

    [void]($global:TabControl.Items.Add($global:AddDepartmentTab))

    Write-Log -Message "XML betöltés kész" -Level "INFO"
    $global:Status.Text = "Betöltve"
}