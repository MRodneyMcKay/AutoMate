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
Import-Module (Join-Path -Path $resolvedModulePath -ChildPath 'GuiStyling\GuiStyling.psd1')

Add-Type -AssemblyName System.Windows.Forms

$script:SelectedSheetPath = $null

$TaskGroups = @(
    @{
        Header = 'Front Office'
        Tasks = @(
            @{ Name = 'Igények - Front Office'; RequiresSheetPath = $false; Action = { Write-Log 'Printing schedule requests for the front office'; Print-RequestFrontOffice } },
            @{ Name = 'Jelenléti ív - Front Office'; RequiresSheetPath = $true; Action = { Write-Log 'Printing attendance sheets for the front office'; Print-AttandanceSheetFrontOffice -OpenFile $script:SelectedSheetPath } }
        )
    },
    @{
        Header = 'Fürdő'
        Tasks = @(
            @{ Name = 'Igények - Fürdő'; RequiresSheetPath = $false; Action = { Write-Log 'Printing schedule requests for the staff'; Print-RequestUszomester } },
            @{ Name = 'Jelenléti ív - Fürdő'; RequiresSheetPath = $true; Action = { Write-Log 'Printing attendance sheets for the staff'; Print-AttandanceSheetUszomester -OpenFile $script:SelectedSheetPath } },
            @{ Name = 'Utazási támogatás'; RequiresSheetPath = $false; Action = { Write-Log 'Printing commuting allowance'; Print-CommutingAllowance } }
        )
    },
    @{
        Header = 'Domb Beach'
        Tasks = @(
            @{ Name = 'Igények - Domb Beach'; RequiresSheetPath = $false; Action = { Write-Log 'Printing schedule requests for Domb Beach'; Print-RequestDombBeach } }
        )
    },
    @{
        Header = 'Karbantartók'
        Tasks = @(
            @{ Name = 'Jelenléti ív - Karbantartó'; RequiresSheetPath = $true; Action = { Write-Log 'Printing attendance sheets for the genitors'; Print-AttandanceSheetKarbantarto -OpenFile $script:SelectedSheetPath } }
        )
    },
    @{
        Header = 'Vízgépész'
        Tasks = @(
            @{ Name = 'Jelenléti ív - Gépész'; RequiresSheetPath = $true; Action = { Write-Log 'Printing attendance sheets for the pool technicians'; Print-AttandanceSheetGepesz -OpenFile $script:SelectedSheetPath } }
        )
    },
    @{
        Header = 'Gyógyászat'
        Tasks = @(
            @{ Name = 'Igények - Gyógyászat'; RequiresSheetPath = $false; Action = { Write-Log 'Printing schedule requests for the medical department'; Print-RequestGyogyaszat } },
            @{ Name = 'Jelenléti ív - Gyógyászat'; RequiresSheetPath = $true; Action = { Write-Log 'Printing attendance sheets for the medical department'; Print-AttandanceSheetGyogyaszat -OpenFile $script:SelectedSheetPath } }
        )
    },
    @{
        Header = 'Önkormányzat'
        Tasks = @(
            @{ Name = 'Jelenléti ív - Önkormányzat'; RequiresSheetPath = $true; Action = { Write-Log 'Printing attendance sheets for the municipality'; Print-AttandanceSheetOnkormanyzat -OpenFile $script:SelectedSheetPath } }
        )
    }
)

[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Print Manager"
        Height="760"
        Width="1100"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        AllowDrop="True"
        FontFamily="Segoe UI"
        FontSize="14"
        Background="{DynamicResource BackgroundBrush}">
    <Window.Resources>
        $(Get-WpfSharedStylesXaml -IncludeButtonStyles -IncludeTextBoxStyles)
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Margin" Value="0,4,0,4"/>
        </Style>
        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="Foreground" Value="{StaticResource PrimaryBrush}"/>
        </Style>
        <Style x:Key="CardBorder" TargetType="Border">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="10"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="Margin" Value="0,0,0,12"/>
        </Style>
    </Window.Resources>

    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Style="{StaticResource CardBorder}">
            <StackPanel>
                <TextBlock Text="Nyomtatási kezelő" FontSize="24" FontWeight="Bold" Margin="0,0,0,6"/>
                <TextBlock Text="Válasszon egy jelenléti ív XLS/XLSX fájlt, majd jelölje ki a nyomtatandó feladatokat." TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}"/>
            </StackPanel>
        </Border>

        <Border Grid.Row="1" Style="{StaticResource CardBorder}">
            <StackPanel>
                <TextBlock Text="Jelenléti ív fájl" Style="{StaticResource SectionHeader}"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox Name="PathBox" Margin="0,0,10,0" IsReadOnly="True" Text="Nincs fájl kiválasztva"/>
                    <Button Name="BrowseButton" Grid.Column="1" Content="Tallózás" Style="{StaticResource PrimaryButton}" Width="110"/>
                </Grid>
                <TextBlock Text="A kiválasztott fájl minden nyomtatási feladathoz használatos." Margin="0,8,0,0" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
            </StackPanel>
        </Border>

        <Border Grid.Row="2" Style="{StaticResource CardBorder}">
            <StackPanel>
                <Grid Margin="0,0,0,8">
                    <TextBlock Text="Feladatok" Style="{StaticResource SectionHeader}"/>
                    <CheckBox Name="SelectAllBox" Content="Összes kijelölése" HorizontalAlignment="Right" VerticalAlignment="Center"/>
                </Grid>
                <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="360">
                    <WrapPanel Name="TaskList" Orientation="Horizontal" HorizontalAlignment="Center"/>
                </ScrollViewer>
            </StackPanel>
        </Border>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button Name="RunButton" Content="Nyomtatás" Style="{StaticResource PrimaryButton}" Width="140" Margin="0,0,10,0"/>
            <Button Name="CloseButton" Content="Bezárás" Width="110"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$Window = [Windows.Markup.XamlReader]::Load($reader)

$TaskList = $Window.FindName('TaskList')
$TaskList.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
$RunButton = $Window.FindName('RunButton')
$SelectAllBox = $Window.FindName('SelectAllBox')
$BrowseButton = $Window.FindName('BrowseButton')
$PathBox = $Window.FindName('PathBox')
$CloseButton = $Window.FindName('CloseButton')

function Set-AttendanceSheetPath {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path $Path)) {
        return
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($extension -notin '.xls', '.xlsx', '.xlsm') {
        [System.Windows.MessageBox]::Show(
            'Csak Excel fájl (*.xls, *.xlsx, *.xlsm) választható ki.',
            'Érvénytelen fájl',
            'OK',
            'Warning'
        ) | Out-Null

        return
    }

    $script:SelectedSheetPath = (Resolve-Path $Path).Path
    $PathBox.Text = $script:SelectedSheetPath
}

$Tasks = @{}
$CheckBoxes = @{}
$TaskMetadata = @{}

foreach ($group in $TaskGroups) {
    $section = New-Object System.Windows.Controls.Border
    $section.Style = $Window.FindResource('CardBorder')
    $section.Width = 240
    $section.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
    $section.Margin = '0,0,12,12'

    $content = New-Object System.Windows.Controls.StackPanel
    $header = New-Object System.Windows.Controls.TextBlock
    $header.Text = $group.Header
    $header.FontWeight = 'Bold'
    $header.Margin = '0,0,0,6'
    $content.Children.Add($header)

    $sep = New-Object System.Windows.Controls.Separator
    $sep.Margin = '0,0,0,8'
    $content.Children.Add($sep)

    foreach ($task in $group.Tasks) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = ($task.Name -split ' - ')[0]
        $cb.ToolTip = if ($task.RequiresSheetPath) { 'Ez a feladat egy kiválasztott jelenléti ív fájlt igényel.' } else { 'Ez a feladat nem igényel jelenléti ív fájlt.' }
        $content.Children.Add($cb)
        $CheckBoxes[$task.Name] = $cb
        $Tasks[$task.Name] = $task.Action
        $TaskMetadata[$task.Name] = $task
    }

    $section.Child = $content
    $TaskList.Children.Add($section)
}

$SelectAllBox.Add_Checked({
    foreach ($cb in $CheckBoxes.Values) { $cb.IsChecked = $true }
})
$SelectAllBox.Add_Unchecked({
    foreach ($cb in $CheckBoxes.Values) { $cb.IsChecked = $false }
})

$BrowseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Jelenléti ív kiválasztása'
    $dialog.Filter = 'Excel files (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm'
    $dialog.InitialDirectory = $env:USERPROFILE

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Set-AttendanceSheetPath $dialog.FileName
    }
})

$Window.Add_DragOver({

    if ($_.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {

        $files = $_.Data.GetData([System.Windows.DataFormats]::FileDrop)

        if ($files.Count -eq 1) {

            $extension = [System.IO.Path]::GetExtension($files[0]).ToLowerInvariant()

            if ($extension -in '.xls', '.xlsx', '.xlsm') {
                $_.Effects = [System.Windows.DragDropEffects]::Copy
            }
            else {
                $_.Effects = [System.Windows.DragDropEffects]::None
            }
        }
        else {
            $_.Effects = [System.Windows.DragDropEffects]::None
        }
    }
    else {
        $_.Effects = [System.Windows.DragDropEffects]::None
    }

    $_.Handled = $true
})

$Window.Add_Drop({

    if ($_.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {

        $files = $_.Data.GetData([System.Windows.DataFormats]::FileDrop)

        if ($files.Count -gt 0) {
            Set-AttendanceSheetPath $files[0]
        }
    }

    $_.Handled = $true
})

$RunButton.Add_Click({
    $selectedTasks = @($Tasks.Keys | Where-Object { $CheckBoxes[$_].IsChecked })
    $requiresSheetPath = @($selectedTasks | Where-Object { $TaskMetadata[$_].RequiresSheetPath })

    if ($requiresSheetPath.Count -gt 0 -and -not $script:SelectedSheetPath) {
        [System.Windows.MessageBox]::Show('A kijelölt nyomtatási feladatokhoz előbb válasszon ki egy jelenléti ív fájlt.', 'Hiányzó fájl', 'OK', 'Warning') | Out-Null
        return
    }

    foreach ($taskName in $selectedTasks) {
        try {
            Write-Log "Running: $taskName"
            & $Tasks[$taskName]
            Write-Log "Done: $taskName"
        }
        catch {
            Write-Log "Error in ${taskName}: $($_.Exception.Message)" -Level Error -ShowMessageBox
        }
    }
})

$CloseButton.Add_Click({ $Window.Close() })

$Window.Topmost = $true
$Window.ShowDialog() | Out-Null
