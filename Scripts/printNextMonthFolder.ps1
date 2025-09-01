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
Add-Type -AssemblyName System.Xaml

# Load modules and get file path
$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$resolvedModulePath = (Resolve-Path -Path $ModulePath).Path

Import-Module (Join-Path -Path $resolvedModulePath -ChildPath 'LoggingSystem\LoggingSystem.psd1')
Import-Module (Join-Path -Path $resolvedModulePath -ChildPath 'PrintNextMonthFolder\PrintNextMonthFolder.psd1')

$path = Get-SheetPath

# Define task groups
$TaskGroups = @(
    @{
        Header = "Front Office"
        Tasks = @{
            "Igények - Front Office" = {
                Write-Log "Printing schedule requests for the front office"
                Print-RequestFrontOffice
            }
            "Jelenléti ív - Front Office" = {
                Write-Log "Printing attendance sheets for the front office"
                Print-AttandanceSheetFrontOffice -OpenFile $path
            }
        }
    },
    @{
        Header = "Fürdő"
        Tasks = @{
            "Igények - Fürdő" = {
                Write-Log "Printing schedule requests for the staff"
                Print-RequestUszomester
            }
            "Jelenléti ív - Fürdő" = {
                Write-Log "Printing attendance sheets for the staff"
                Print-AttandanceSheetUszomester -OpenFile $path
            }
            "Utazási támogatás" = {
                Write-Log "Printing commuting allowance"
                Print-CommutingAllowance
            }            
        }
    },
    @{
        Header = "Karbantartók"
        Tasks = @{
            "Jelenléti ív - Karbantartó" = {
                Write-Log "Printing attendance sheets for the genitors"
                Print-AttandanceSheetKarbantarto -OpenFile $path
            }
        }
    },
    @{
        Header = "Vízgépész"
        Tasks = @{
            "Jelenléti ív - Gépész" = {
                Write-Log "Printing attendance sheets for the pool technicians"
                Print-AttandanceSheetGepesz -OpenFile $path
            }
        }
    },
    @{
        Header = "Gyógyászat"
        Tasks = @{
            "Jelenléti ív - Gyógyászat" = {
                Write-Log "Printing attendance sheets for the mediacal department"
                Print-AttandanceSheetGyogyaszat -OpenFile $path
            }
        }
    }
)

# XAML layout
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Print Manager" Height="550" Width="440" WindowStartupLocation="CenterScreen">
    <DockPanel Margin="10">
        <StackPanel  DockPanel.Dock="Top">
            <CheckBox Name="SelectAllBox" Content="Select All" Margin="0,0,0,10" />
            <ScrollViewer VerticalScrollBarVisibility="Disabled">
                <StackPanel Name="TaskList" />
            </ScrollViewer>
        </StackPanel>

        <Button DockPanel.Dock="Top" Name="RunButton" Height="40" Margin="0,10,0,10" Content="Nyomtatás" />
    </DockPanel>
</Window>
"@

# Load XAML and connect elements
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$Window = [Windows.Markup.XamlReader]::Load($reader)

$TaskList = $Window.FindName("TaskList")
$RunButton = $Window.FindName("RunButton")
$SelectAllBox = $Window.FindName("SelectAllBox")

# Store all tasks and checkboxes
$Tasks = @{}
$CheckBoxes = @{}

# Add UI elements by section
foreach ($group in $TaskGroups) {
    # Header
    $header = New-Object System.Windows.Controls.TextBlock
    $header.Text = $group.Header
    $header.FontWeight = 'Bold'
    $header.Margin = '0,10,0,2'
    $TaskList.Children.Add($header)

    # Separator
    $sep = New-Object System.Windows.Controls.Separator
    $sep.Margin = '0,0,0,10'
    $TaskList.Children.Add($sep)

    # Task checkboxes
    foreach ($taskName in $group.Tasks.Keys) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = $taskName
        $cb.Margin = '0,5,0,5'
        $TaskList.Children.Add($cb)
        $CheckBoxes[$taskName] = $cb
        $Tasks[$taskName] = $group.Tasks[$taskName]
    }
}

# Select All checkbox behavior
$SelectAllBox.Add_Checked({
    foreach ($cb in $CheckBoxes.Values) { $cb.IsChecked = $true }
})
$SelectAllBox.Add_Unchecked({
    foreach ($cb in $CheckBoxes.Values) { $cb.IsChecked = $false }
})

# Run button behavior
$RunButton.Add_Click({
    foreach ($taskName in $Tasks.Keys) {
        if ($CheckBoxes[$taskName].IsChecked) {
            try {
                Write-Log "Running: ${taskName}"
                & $Tasks[$taskName]
                Write-Log "Done: ${taskName}"
            } catch {
                Write-Log "Error in ${taskName}: $($_.Exception.Message)" -Level Error
            }
        }
    }
})

# Show window
$Window.Topmost = $true
$Window.ShowDialog() | Out-Null
