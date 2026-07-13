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

$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$resolvedModulegPath = (Resolve-Path -Path $ModulePath).Path
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'LoggingSystem\LoggingSystem.psd1')
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'GuiStyling\GuiStyling.psd1') -Force

# Environment variable
$EnvVarName = "EfoNévsor"
$cultureHU = [System.Globalization.CultureInfo]::GetCultureInfo("hu-HU")

# Read environment variable
function Read-EfoNames {
    $raw = [Environment]::GetEnvironmentVariable($EnvVarName, "User")
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    $raw -split ';' | Where-Object { $_ } | Sort-Object -Culture $cultureHU
}

# Write environment variable
function Write-EfoNames {
    param([string[]]$Names)
    try {
        $value = ($Names | Sort-Object -Culture $cultureHU) -join ';'
        [Environment]::SetEnvironmentVariable($EnvVarName, $value, "User")
        return $true
    } catch {
        return $false
    }
}

# ObservableCollection backing the list
$Names = New-Object System.Collections.ObjectModel.ObservableCollection[string]
foreach ($n in (Read-EfoNames)) { $Names.Add($n) }

function Sort-EfoNames {
    $Sorted = $Names | Sort-Object -Culture $cultureHU
    $Names.Clear()
    foreach ($n in $Sorted) { $Names.Add($n) }
}

# --- DIRTY STATE TRACKING ---
$global:IsDirty = $false

function Update-DirtyUI {
    if ($global:IsDirty) {
        $ReloadButton.Visibility = "Visible"
        $Status.Text = "Vannak mentetlen módosítások"
    }
    else {
        $ReloadButton.Visibility = "Collapsed"
        $Status.Text = "Készen áll"
    }
}

# Theme + XAML
$Theme = Get-ThemePalette
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="EFO Névsor"
        Width="620" Height="550"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="14"
        Background="$($Theme.BackgroundHex)">

    <Window.Resources>
        $(Get-WpfSharedStylesXaml -IncludeButtonStyles -IncludeTextBoxStyles -IncludeListBoxStyles)
    </Window.Resources>

    <DockPanel>
        <Border DockPanel.Dock="Top" Background="{StaticResource SurfaceBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,0,0,1" Padding="20,15">
            <DockPanel>
                <TextBlock Text="EFO Névsor" FontSize="20" FontWeight="Bold" Foreground="{StaticResource TextBrush}" VerticalAlignment="Center" DockPanel.Dock="Left"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" DockPanel.Dock="Right">
                    <Button Name="ReloadButton" Content="Módosítások elvetése" Width="180" Margin="0,0,10,0" Visibility="Collapsed"/>
                    <Button Name="SaveButton" Content="Mentés" Width="120" Style="{StaticResource PrimaryButton}"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <Border DockPanel.Dock="Bottom" Background="{StaticResource SurfaceBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,1,0,0" Padding="20,5">
            <TextBlock Name="Status" Text="Készen áll" FontSize="12" Foreground="{StaticResource MutedBrush}"/>
        </Border>

        <Border Padding="20">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <ListBox Name="NameList" Grid.Row="0" Margin="0,0,0,15"/>

                <StackPanel Grid.Row="1" Orientation="Horizontal">
                    <TextBox Name="NameBox"
                             Width="260"
                             AcceptsReturn="True"
                             TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto"
                             VerticalContentAlignment="Top"/>
                    <Button Name="CancelButton" Content="Mégse" Width="90" Margin="10,0,0,0" Visibility="Collapsed"/>
                    <Button Name="AddButton" Content="Hozzáad" Width="100" Margin="10,0,0,0" Style="{StaticResource PrimaryButton}"/>
                    <Button Name="DeleteButton" Content="Törlés" Width="90" Margin="10,0,0,0" Visibility="Collapsed"/>
                </StackPanel>
            </Grid>
        </Border>
    </DockPanel>
</Window>
"@

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# UI references
$NameList     = $window.FindName("NameList")
$NameBox      = $window.FindName("NameBox")
$AddButton    = $window.FindName("AddButton")
$DeleteButton = $window.FindName("DeleteButton")
$CancelButton = $window.FindName("CancelButton")
$SaveButton   = $window.FindName("SaveButton")
$ReloadButton = $window.FindName("ReloadButton")
$Status       = $window.FindName("Status")

$NameList.ItemsSource = $Names

# Add a brand-new name
function Add-EfoName {
    $NewName = $NameBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($NewName)) { return }

    if ($Names | Where-Object { $_.Equals($NewName, 'InvariantCultureIgnoreCase') }) {
        Write-Log -Message "Duplikált név: $NewName" -Level "WARNING"
        $Status.Text = "Duplikált név"
        return
    }

    $Names.Add($NewName)
    Sort-EfoNames

    $global:IsDirty = $true
    Update-DirtyUI

    Write-Log -Message "Név hozzáadva: $NewName" -Level "INFO"
    $Status.Text = "Név hozzáadva: $NewName"
    $NameBox.Clear()
}

# Rename the currently selected name
function Update-EfoName {
    $OldName = $NameList.SelectedItem
    if ($null -eq $OldName) {
        $Status.Text = "Nincs kiválasztott név"
        return
    }

    $NewName = $NameBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($NewName)) {
        $Status.Text = "A név nem lehet üres"
        return
    }

    if ($NewName -eq $OldName) {
        $NameBox.Clear()
        $NameList.SelectedItem = $null
        $Status.Text = "Nincs változás"
        return
    }

    if ($Names | Where-Object { $_.Equals($NewName, 'InvariantCultureIgnoreCase') }) {
        Write-Log -Message "Frissítés sikertelen, duplikált név: $NewName" -Level "WARNING"
        $Status.Text = "Ez a név már létezik"
        return
    }

    $Index = $Names.IndexOf($OldName)
    if ($Index -lt 0) {
        $Status.Text = "A név már nem található"
        return
    }

    $Names[$Index] = $NewName
    Sort-EfoNames

    $global:IsDirty = $true
    Update-DirtyUI

    Write-Log -Message "Név módosítva: $OldName -> $NewName" -Level "INFO"
    $Status.Text = "Név módosítva: $NewName"

    $NameBox.Clear()
    $NameList.SelectedItem = $null
}

# Remove the currently selected name
function Remove-EfoName {
    $Selected = $NameList.SelectedItem
    if ($null -ne $Selected) {
        $Names.Remove($Selected)

        $global:IsDirty = $true
        Update-DirtyUI

        Write-Log -Message "Név törölve: $Selected" -Level "INFO"
        $Status.Text = "Név törölve"
        $NameBox.Clear()
    }
    else {
        $Status.Text = "Nincs kiválasztott név"
    }
}

# Selecting a name in the list switches the row into edit mode
$NameList.Add_SelectionChanged({
    param($Sender, $Event)
    if ($null -ne $NameList.SelectedItem) {
        $NameBox.Text = $NameList.SelectedItem
        $AddButton.Content = "Frissítés"
        $DeleteButton.Visibility = "Visible"
    }
    else {
        $AddButton.Content = "Hozzáad"
        $DeleteButton.Visibility = "Collapsed"
    }
})

# Text box housekeeping: strip stray line breaks, show/hide Mégse
$NameBox.Add_TextChanged({
    param($sender, $event)

    $trimmed = $sender.Text.TrimEnd("`r", "`n")
    if ($trimmed -ne $sender.Text) {
        $caret = $trimmed.Length
        $sender.Text = $trimmed
        $sender.CaretIndex = $caret
    }

    if ([string]::IsNullOrWhiteSpace($sender.Text)) {
        $CancelButton.Visibility = "Collapsed"
    }
    else {
        $CancelButton.Visibility = "Visible"
    }
})

$AddButton.Add_Click({
    if ($null -ne $NameList.SelectedItem) {
        Update-EfoName
    }
    else {
        Add-EfoName
    }
})

$DeleteButton.Add_Click({ Remove-EfoName })

$CancelButton.Add_Click({
    $NameBox.Clear()
    $NameList.SelectedItem = $null
})

$NameBox.Add_KeyDown({
    param($Sender, $Event)
    if ($Event.Key -eq "Enter") {
        if ($null -ne $NameList.SelectedItem) {
            Update-EfoName
        }
        else {
            Add-EfoName
        }
        $Event.Handled = $true
    }
})

# Explicit save / discard
$SaveButton.Add_Click({
    try {
        if (Write-EfoNames -Names $Names) {
            Write-Log -Message "EFO névsor mentve" -Level "INFO"
            $Status.Text = "Mentés kész"

            $global:IsDirty = $false
            Update-DirtyUI
        }
        else {
            Write-Log -Message "Hiba a névsor mentése közben" -Level "ERROR" -ShowMessageBox
            $Status.Text = "Mentés sikertelen"
        }
    }
    catch {
        Write-Log -Message "Hiba a névsor mentése közben: $($_.Exception.Message)" -Level "ERROR" -ShowMessageBox
        $Status.Text = "Mentés sikertelen"
    }
})

$ReloadButton.Add_Click({
    $Names.Clear()
    foreach ($n in (Read-EfoNames)) { $Names.Add($n) }
    $NameBox.Clear()
    $NameList.SelectedItem = $null

    $global:IsDirty = $false
    Update-DirtyUI

    Write-Log -Message "EFO névsor újratöltve" -Level "INFO"
    $Status.Text = "Módosítások elvetve"
})

# Show window
$window.ShowDialog() | Out-Null
