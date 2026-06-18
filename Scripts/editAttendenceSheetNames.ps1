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

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase


# Module imports

$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$ResolvedModulePath = (Resolve-Path -Path $ModulePath).Path

Import-Module (Join-Path -Path $ResolvedModulePath -ChildPath "LoggingSystem\LoggingSystem.psd1")


# Configuration

$CsvFolder = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények"

$CultureHU = [System.Globalization.CultureInfo]::GetCultureInfo("hu-HU")


# State

$Tabs = @{}
$CsvFiles = @{}
$TabControls = @{}
$DirtyTabs = New-Object System.Collections.Generic.HashSet[string]


Write-Log -Message "Jelenléti szerkesztő indul" -Level "INFO"


# Sort names alphabetically

function Sort-Names {
    param (
        [System.Collections.ObjectModel.ObservableCollection[string]]$Collection
    )

    $Sorted = $Collection | Sort-Object -Culture $CultureHU

    $Collection.Clear()

    foreach ($Name in $Sorted) {
        $Collection.Add($Name)
    }
}


# Mark tab as modified

function Mark-Dirty {
    param (
        [string]$TabName
    )

    if ($DirtyTabs.Add($TabName)) {

        if ($TabControls.ContainsKey($TabName)) {

            $TabControls[$TabName].Tab.Header = "*$TabName"
        }
    }
}


# Clear modified state

function Clear-Dirty {
    param (
        [string]$TabName
    )

    $DirtyTabs.Remove($TabName) | Out-Null

    if ($TabControls.ContainsKey($TabName)) {

        $TabControls[$TabName].Tab.Header = $TabName
    }
}


# Save CSV files

function Save-All {

    foreach ($TabName in $Tabs.Keys) {

        try {

            $Lines = New-Object System.Collections.Generic.List[string]

            $Lines.Add("Név")


            foreach ($Name in ($Tabs[$TabName] | Sort-Object -Culture $CultureHU)) {

                $Lines.Add($Name)
            }


            [System.IO.File]::WriteAllLines(
                $CsvFiles[$TabName],
                $Lines,
                [System.Text.UTF8Encoding]::new($true)
            )


            Clear-Dirty -TabName $TabName

            Write-Log -Message "$TabName mentve" -Level "INFO"

        }
        catch {

            Write-Log `
                -Message "Mentési hiba: $TabName $($_.Exception.Message)" `
                -Level "ERROR" `
                -ShowMessageBox
        }
    }


    $Status.Text = "Mentés kész"
}

# Load CSV files

function Load-All {

    $TabControl.Items.Clear()

    $Tabs.Clear()
    $CsvFiles.Clear()
    $TabControls.Clear()
    $DirtyTabs.Clear()


    if (-not (Test-Path -Path $CsvFolder)) {

        Write-Log `
            -Message "Nem található: $CsvFolder" `
            -Level "ERROR" `
            -ShowMessageBox

        return
    }


    $Files = Get-ChildItem `
        -Path $CsvFolder `
        -Filter "*.csv" `
        -Force |
        Sort-Object Name


    Write-Log -Message "$($Files.Count) CSV található" -Level "INFO"


    foreach ($File in $Files) {

        $TabName = [IO.Path]::GetFileNameWithoutExtension($File.Name)

        Write-Log -Message "Betöltés: $TabName" -Level "INFO"


        $Collection = New-Object System.Collections.ObjectModel.ObservableCollection[string]


        foreach ($Row in Import-Csv -Path $File.FullName) {

            if ($Row.Név) {

                $Collection.Add($Row.Név.Trim())
            }
        }


        Sort-Names -Collection $Collection


        $Tabs[$TabName] = $Collection
        $CsvFiles[$TabName] = $File.FullName


        # Create tab

        $Tab = New-Object System.Windows.Controls.TabItem

        $Tab.Header = $TabName


        $Grid = New-Object System.Windows.Controls.Grid


        $ListRow = New-Object Windows.Controls.RowDefinition
        $ListRow.Height = "*"


        $InputRow = New-Object Windows.Controls.RowDefinition
        $InputRow.Height = "Auto"


        $Grid.RowDefinitions.Add($ListRow)
        $Grid.RowDefinitions.Add($InputRow)


        $List = New-Object System.Windows.Controls.ListBox

        $List.Margin = "5"
        $List.ItemsSource = $Collection


        [Windows.Controls.Grid]::SetRow($List, 0)


        $Panel = New-Object System.Windows.Controls.StackPanel

        $Panel.Orientation = "Horizontal"
        $Panel.Margin = "5"


        $Box = New-Object System.Windows.Controls.TextBox

        $Box.Width = 250
        $Box.Height = 25
        $Box.VerticalContentAlignment = "Center"


        $AddButton = New-Object System.Windows.Controls.Button

        $AddButton.Width = 90
        $AddButton.Height = 25
        $AddButton.Content = "Hozzáad"


        $DeleteButton = New-Object System.Windows.Controls.Button

        $DeleteButton.Width = 90
        $DeleteButton.Height = 25
        $DeleteButton.Content = "Törlés"
        $DeleteButton.Margin = "5,0,0,0"


        $Panel.Children.Add($Box)
        $Panel.Children.Add($AddButton)
        $Panel.Children.Add($DeleteButton)


        [Windows.Controls.Grid]::SetRow($Panel, 1)


        $Grid.Children.Add($List)
        $Grid.Children.Add($Panel)


        $Tab.Content = $Grid


        $TabControl.Items.Add($Tab)


        $TabControls[$TabName] = @{
            Box          = $Box
            Button       = $AddButton
            DeleteButton = $DeleteButton
            Collection   = $Collection
            List         = $List
            Tab          = $Tab
        }


        # Add button

        $AddButton.Add_Click({

            $CurrentTab = $TabControl.SelectedItem

            if ($null -eq $CurrentTab) {
                return
            }


            $CurrentName = $CurrentTab.Header.ToString().TrimStart("*")

            $Data = $TabControls[$CurrentName]

            $NewName = $Data.Box.Text.Trim()


            if ([string]::IsNullOrWhiteSpace($NewName)) {
                return
            }


            if ($Data.Collection -contains $NewName) {

                Write-Log `
                    -Message "Duplikált név: $NewName" `
                    -Level "WARNING"

                return
            }


            $Data.Collection.Add($NewName)

            Sort-Names -Collection $Data.Collection


            Mark-Dirty -TabName $CurrentName


            $Data.Box.Clear()


            Write-Log `
                -Message "Név hozzáadva: $NewName" `
                -Level "INFO"


            $Status.Text = "Név hozzáadva"

        })


        # Delete button

        $DeleteButton.Add_Click({

            $CurrentTab = $TabControl.SelectedItem

            if ($null -eq $CurrentTab) {
                return
            }


            $CurrentName = $CurrentTab.Header.ToString().TrimStart("*")

            $Data = $TabControls[$CurrentName]


            $Selected = $Data.List.SelectedItem


            if ($Selected) {

                $Data.Collection.Remove($Selected)


                Mark-Dirty -TabName $CurrentName


                Write-Log `
                    -Message "Törölve: $Selected" `
                    -Level "INFO"


                $Status.Text = "Törölve"

            }
            else {

                $Status.Text = "Nincs kiválasztott név"
            }

        })

                # Enter key handler

        $Box.Add_KeyDown({

            param (
                $Sender,
                $Event
            )


            if ($Event.Key -eq "Enter") {

                $CurrentTab = $TabControl.SelectedItem

                if ($null -eq $CurrentTab) {
                    return
                }


                $CurrentName = $CurrentTab.Header.ToString().TrimStart("*")

                $Data = $TabControls[$CurrentName]


                $NewName = $Data.Box.Text.Trim()


                if (-not [string]::IsNullOrWhiteSpace($NewName)) {


                    if (-not ($Data.Collection -contains $NewName)) {


                        $Data.Collection.Add($NewName)

                        Sort-Names -Collection $Data.Collection


                        Mark-Dirty -TabName $CurrentName


                        $Data.Box.Clear()


                        Write-Log `
                            -Message "Név hozzáadva: $NewName" `
                            -Level "INFO"


                        $Status.Text = "Név hozzáadva"
                    }
                }


                $Event.Handled = $true
            }

        })

    }


    Write-Log -Message "CSV betöltés kész" -Level "INFO"

    $Status.Text = "Betöltve"
}


# XAML

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Jelenléti szerkesztő"
        Width="800"
        Height="600"
        WindowStartupLocation="CenterScreen">

    <DockPanel Margin="10">

        <ToolBar DockPanel.Dock="Top">

            <Button Name="SaveButton"
                    Content="Mentés"
                    Width="100"/>

            <Button Name="ReloadButton"
                    Content="Módosítások elvetése"
                    Width="160"
                    Margin="10,0,0,0"/>

        </ToolBar>


        <StatusBar DockPanel.Dock="Bottom">

            <TextBlock Name="Status"
                       Text="Készen áll"/>

        </StatusBar>


        <TabControl Name="TabControl"/>

    </DockPanel>

</Window>
"@


# Create window

$Reader = New-Object System.Xml.XmlNodeReader ([xml]$Xaml)

$Window = [Windows.Markup.XamlReader]::Load($Reader)


$SaveButton = $Window.FindName("SaveButton")
$ReloadButton = $Window.FindName("ReloadButton")
$TabControl = $Window.FindName("TabControl")
$Status = $Window.FindName("Status")


# Button events

$SaveButton.Add_Click({

    Save-All

})


$ReloadButton.Add_Click({

    Load-All

})


# Start application

Load-All

$Window.ShowDialog() | Out-Null