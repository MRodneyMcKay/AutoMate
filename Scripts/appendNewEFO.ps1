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

# Import your logging system
$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$resolvedModulegPath = (Resolve-Path -Path $ModulePath).Path
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'LoggingSystem\LoggingSystem.psd1')

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

# ObservableCollection
$Names = New-Object System.Collections.ObjectModel.ObservableCollection[System.String]
foreach ($n in (Read-EfoNames)) { $Names.Add($n) }

# Cancel flag
$CancelClicked = $false

# XAML
$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        Title='EFO Névsor' Height='350' Width='600'
        WindowStartupLocation='CenterScreen'>
    <DockPanel Margin='10'>
        <StatusBar DockPanel.Dock='Bottom'>
            <TextBlock Name='StatusBar' Text='Készen áll' />
        </StatusBar>
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width='2*'/>
                <ColumnDefinition Width='Auto'/>
                <ColumnDefinition Width='3*'/>
            </Grid.ColumnDefinitions>

            <!-- Left: Input -->
            <Grid Grid.Column='0'>
                <Grid.RowDefinitions>
                    <RowDefinition Height='*'/>
                    <RowDefinition Height='Auto'/>
                    <RowDefinition Height='*'/>
                </Grid.RowDefinitions>
                <TextBox Grid.Row='1'
                         Name='NameBox'
                         Width='180'
                         Height='28'
                         VerticalContentAlignment='Center'
                         Tag='Név'
                         Foreground='Gray'/>
            </Grid>

            <!-- Middle: Add, Trash, Cancel -->
            <Grid Grid.Column='1' Name='MiddleGrid' Margin='15,0'>
                <Grid.RowDefinitions>
                    <RowDefinition Height='*'/>
                    <RowDefinition Height='Auto'/>
                    <RowDefinition Height='Auto'/>
                    <RowDefinition Height='Auto'/>
                    <RowDefinition Height='*'/>
                </Grid.RowDefinitions>

                <!-- Add Arrow -->
                <Button Grid.Row='1'
                        Name='AddButton'
                        Width='40'
                        Height='40'
                        FontFamily='Segoe MDL2 Assets'
                        FontSize='20'
                        Content='&#xE72A;' />

                <!-- Trash -->
                <Button Grid.Row='2'
                        Name='TrashButton'
                        Width='40'
                        Height='40'
                        Margin='0,10,0,0'
                        FontFamily='Segoe MDL2 Assets'
                        FontSize='20'
                        Content='&#xE74D;' />

                <!-- Mégse -->
                <Button Grid.Row='3'
                        Name='CancelButton'
                        Width='60'
                        Height='30'
                        Margin='0,10,0,0'
                        Content='Mégse' />
            </Grid>

            <!-- Right: List of names -->
            <Grid Grid.Column='2'>
                <ListBox Name='NameList' />
            </Grid>
        </Grid>
    </DockPanel>
</Window>
"@

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# UI references
$NameList    = $window.FindName("NameList")
$NameBox     = $window.FindName("NameBox")
$AddButton   = $window.FindName("AddButton")
$TrashButton = $window.FindName("TrashButton")
$CancelButton= $window.FindName("CancelButton")
$StatusBar   = $window.FindName("StatusBar")

# Bind list
$NameList.ItemsSource = $Names

# Placeholder logic
$NameBox.Text = $NameBox.Tag
$NameBox.Foreground = "Gray"
$NameBox.Add_GotFocus({ if ($NameBox.Text -eq $NameBox.Tag) { $NameBox.Text = ""; $NameBox.Foreground="Black" } })
$NameBox.Add_LostFocus({ if ([string]::IsNullOrWhiteSpace($NameBox.Text)) { $NameBox.Text=$NameBox.Tag; $NameBox.Foreground="Gray" } })

# Add name
function Add-Name {
    $newName = $NameBox.Text.Trim()
    if (-not $newName -or $newName -eq $NameBox.Tag) { return }
    if ($Names | Where-Object { $_.Equals($newName,'InvariantCultureIgnoreCase') }) {
        Write-Log -Message "'$newName' már szerepel a névsorban" -Level WARNING -ShowMessageBox
        $StatusBar.Text = "Duplikált név"; return
    }
    $Names.Add($newName)
    # Re-sort
    $sorted = $Names | Sort-Object -Culture $cultureHU
    $Names.Clear()
    foreach ($n in $sorted) { $Names.Add($n) }
    $StatusBar.Text = "Név hozzáadva"
    $NameBox.Text = $NameBox.Tag
    $NameBox.Foreground = "Gray"
}

$AddButton.Add_Click({ Add-Name })
$NameBox.Add_KeyDown({ param($s,$e); if ($e.Key -eq 'Enter') { Add-Name; $e.Handled=$true } })

# Trash button
$TrashButton.Add_Click({
    $selected = $NameList.SelectedItem
    if ($null -ne $selected) {
        $StatusBar.Text = "Név törlése..."
        $Names.Remove($selected)
        $StatusBar.Text = "Név törölve"
    } else { $StatusBar.Text = "Nincs kiválasztott név" }
})

# Cancel button
$CancelButton.Add_Click({
    $CancelClicked = $true
    $window.Close()
})

# Save on closing (if not canceled)
$window.Add_Closing({
    if (-not $CancelClicked) {
        $StatusBar.Text = "Mentés..."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
        try {
            if (-not (Write-EfoNames -Names $Names)) {
                Write-Log -Message "Hiba a névsor mentése közben ablak bezáráskor" -Level ERROR -ShowMessageBox
            } else { $StatusBar.Text = "Mentés kész" }
        } catch {
            Write-Log -Message "Hiba a névsor mentése közben ablak bezáráskor" -Level ERROR -ShowMessageBox
        }
    }
})

# Show window
$window.ShowDialog() | Out-Null