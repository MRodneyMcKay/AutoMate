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

function Mark-Dirty {
    param (
        [string]$DeptName,
        [string]$PosName = $null
    )

    if ($global:DirtyDepartments.Add($DeptName)) {
        if ($global:DeptControls.ContainsKey($DeptName)) {
            Set-HeaderDirtyState -HeaderBlock $global:DeptControls[$DeptName].HeaderBlock -Name $DeptName -IsDirty $true
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
}

function Clear-AllDirty {
    foreach ($DeptName in $global:DirtyDepartments) {
        if ($global:DeptControls.ContainsKey($DeptName)) {
            Set-HeaderDirtyState -HeaderBlock $global:DeptControls[$DeptName].HeaderBlock -Name $DeptName -IsDirty $false
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
    $PosName = $Context.PosName
    $Controls = $global:DeptControls[$DeptName].PositionControls[$PosName]
    $NewName = $Controls.Box.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($NewName)) { return }

    $Collection = $global:Data[$DeptName][$PosName]

    if ($Collection -contains $NewName) {
        Write-Log -Message "Duplikált név: $DeptName / $PosName / $NewName" -Level "WARNING"
        $global:Status.Text = "Duplikált név"
        return
    }

    $Collection.Add($NewName)
    Sort-Names -Collection $Collection
    Mark-Dirty -DeptName $DeptName -PosName $PosName
    $Controls.Box.Clear()

    Write-Log -Message "Név hozzáadva: $DeptName / $PosName / $NewName" -Level "INFO"
    $global:Status.Text = "Név hozzáadva"
}

function Remove-NameFromCurrentPosition {
    $Context = Get-CurrentDeptAndPosition
    if ($null -eq $Context) {
        $global:Status.Text = "Nincs kiválasztott pozíció"
        return
    }

    $DeptName = $Context.DeptName
    $PosName = $Context.PosName
    $Controls = $global:DeptControls[$DeptName].PositionControls[$PosName]
    $Selected = $Controls.List.SelectedItem

    if ($Selected) {
        $global:Data[$DeptName][$PosName].Remove($Selected)
        Mark-Dirty -DeptName $DeptName -PosName $PosName
        Write-Log -Message "Törölve: $DeptName / $PosName / $Selected" -Level "INFO"
        $global:Status.Text = "Törölve"
    }
    else {
        $global:Status.Text = "Nincs kiválasztott név"
    }
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

    $Grid = New-Object System.Windows.Controls.Grid
    $Grid.Margin = "0,10,0,0"
    
    $ListRow = New-Object Windows.Controls.RowDefinition
    $ListRow.Height = "*"
    $InputRow = New-Object Windows.Controls.RowDefinition
    $InputRow.Height = "Auto"

    $Grid.RowDefinitions.Add($ListRow)
    $Grid.RowDefinitions.Add($InputRow)

    $List = New-Object System.Windows.Controls.ListBox
    $List.Margin = "0,0,0,15"
    $List.ItemsSource = $global:Data[$DeptName][$PosName]
    [Windows.Controls.Grid]::SetRow($List, 0)

    $Panel = New-Object System.Windows.Controls.StackPanel
    $Panel.Orientation = "Horizontal"

    $Box = New-Object System.Windows.Controls.TextBox
    $Box.Width = 300
    $Box.VerticalContentAlignment = "Center"
    
    $AddButton = New-Object System.Windows.Controls.Button
    $AddButton.Content = "Hozzáad"
    $AddButton.Width = 100
    $AddButton.Margin = "10,0,0,0"
    $AddButton.Style = $global:Window.FindResource("PrimaryButton")

    $DeleteButton = New-Object System.Windows.Controls.Button
    $DeleteButton.Content = "Törlés"
    $DeleteButton.Width = 100
    $DeleteButton.Margin = "10,0,0,0"

    $Panel.Children.Add($Box)
    $Panel.Children.Add($AddButton)
    $Panel.Children.Add($DeleteButton)
    [Windows.Controls.Grid]::SetRow($Panel, 1)

    $Grid.Children.Add($List)
    $Grid.Children.Add($Panel)
    $Tab.Content = $Grid

    $PositionTabControl.Items.Add($Tab)

    $global:DeptControls[$DeptName].PositionControls[$PosName] = @{
        Tab          = $Tab
        HeaderBlock  = $HeaderBlock
        Box          = $Box
        AddButton    = $AddButton
        DeleteButton = $DeleteButton
        List         = $List
    }

    $AddButton.Add_Click({ Add-NameToCurrentPosition })
    $DeleteButton.Add_Click({ Remove-NameFromCurrentPosition })

    $Box.Add_KeyDown({
        param ($Sender, $Event)
        if ($Event.Key -eq "Enter") {
            Add-NameToCurrentPosition
            $Event.Handled = $true
        }
    })
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

    if ($global:Data[$DeptName].Contains($PosName)) {
        Write-Log -Message "A pozíció már létezik: $DeptName / $PosName" -Level "WARNING"
        $global:Status.Text = "A pozíció már létezik"
        return
    }

    $global:Data[$DeptName][$PosName] = New-Object System.Collections.ObjectModel.ObservableCollection[string]
    New-PositionTab -DeptName $DeptName -PosName $PosName
    Mark-Dirty -DeptName $DeptName -PosName $PosName

    $global:DeptControls[$DeptName].PositionTabControl.SelectedItem = $global:DeptControls[$DeptName].PositionControls[$PosName].Tab
    Write-Log -Message "Új pozíció hozzáadva: $DeptName / $PosName" -Level "INFO"
    $global:Status.Text = "Pozíció hozzáadva: $PosName"
}

function New-DepartmentTab {
    param (
        [string]$DeptName
    )

    $Tab = New-Object System.Windows.Controls.TabItem
    $Tab.Tag = $DeptName
    $HeaderBlock = New-TabHeaderBlock -Name $DeptName
    $Tab.Header = $HeaderBlock

    $Container = New-Object System.Windows.Controls.Border
    $Container.Background = $global:Window.FindResource("SurfaceBrush")
    $Container.BorderBrush = $global:Window.FindResource("BorderBrush")
    $Container.BorderThickness = "1"
    $Container.CornerRadius = "8"
    $Container.Padding = "15"
    $Container.Margin = "0,10,0,0"

    $DockPanel = New-Object System.Windows.Controls.DockPanel

    $Toolbar = New-Object System.Windows.Controls.StackPanel
    $Toolbar.Orientation = "Horizontal"
    $Toolbar.Margin = "0,0,0,15"
    [Windows.Controls.DockPanel]::SetDock($Toolbar, "Top")

    $AddPositionButton = New-Object System.Windows.Controls.Button
    $AddPositionButton.Content = "+ Új pozíció"
    $AddPositionButton.Width = 120
    $Toolbar.Children.Add($AddPositionButton)

    $PositionTabControl = New-Object System.Windows.Controls.TabControl

    $DockPanel.Children.Add($Toolbar)
    $DockPanel.Children.Add($PositionTabControl)
    $Container.Child = $DockPanel
    $Tab.Content = $Container

    $global:TabControl.Items.Add($Tab)

    $global:DeptControls[$DeptName] = @{
        Tab                = $Tab
        HeaderBlock        = $HeaderBlock
        PositionTabControl = $PositionTabControl
        PositionControls   = @{}
    }

    $AddPositionButton.Add_Click({ Add-Position })
}

function Add-Department {
    $Name = Show-ModernInputBox -Title "Új részleg" -Prompt "Add meg az új részleg nevét:"

    if ([string]::IsNullOrWhiteSpace($Name)) { return }

    if ($global:Data.Contains($Name)) {
        Write-Log -Message "A részleg már létezik: $Name" -Level "WARNING"
        $global:Status.Text = "A részleg már létezik"
        return
    }

    $global:Data[$Name] = [ordered]@{}
    New-DepartmentTab -DeptName $Name
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

            foreach ($PosName in $global:Data[$DeptName].Keys) {
                $PosElement = [System.Xml.Linq.XElement]::new([System.Xml.Linq.XName]"Position")
                $PosElement.SetAttributeValue("Name", $PosName)

                $SortedNames = $global:Data[$DeptName][$PosName] | Sort-Object -Culture $global:CultureHU

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
        Write-Log -Message "Betöltés: $DeptName" -Level "INFO"
        
        $global:Data[$DeptName] = [ordered]@{}
        New-DepartmentTab -DeptName $DeptName

        foreach ($PosElement in $DeptElement.Elements("Position")) {
            $PosName = $PosElement.Attribute("Name").Value
            $Collection = New-Object System.Collections.ObjectModel.ObservableCollection[string]

            foreach ($NevElement in $PosElement.Elements("Nev")) {
                $NameValue = $NevElement.Value.Trim()
                if ($NameValue) { $Collection.Add($NameValue) }
            }

            Sort-Names -Collection $Collection
            $global:Data[$DeptName][$PosName] = $Collection
            New-PositionTab -DeptName $DeptName -PosName $PosName
        }
    }

    Write-Log -Message "XML betöltés kész" -Level "INFO"
    $global:Status.Text = "Betöltve"
}
