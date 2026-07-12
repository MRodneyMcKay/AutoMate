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

$GuiStylingModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\GuiStyling\GuiStyling.psd1"
Import-Module $GuiStylingModulePath -Force

function Get-TabHeaderText {
    param (
        [string]$Name,
        [string]$Facility = ''
    )

    if ([string]::IsNullOrWhiteSpace($Facility) -or $Facility -eq 'Kecskeméti Fürdő') {
        return $Name
    }

    return "$Name - $Facility"
}

function New-TabHeaderBlock {
    param (
        [string]$Name,
        [string]$Facility = ''
    )

    $Theme = Get-ThemePalette
    $Block = New-Object System.Windows.Controls.TextBlock
    $Block.Text = Get-TabHeaderText -Name $Name -Facility $Facility
    $Block.FontSize = 14
    $Block.FontWeight = "SemiBold"
    $Block.Foreground = $Theme.TextHex
    return $Block
}

function Set-HeaderDirtyState {
    param (
        [System.Windows.Controls.TextBlock]$HeaderBlock,
        [string]$Name,
        [string]$Facility = '',
        [bool]$IsDirty
    )

    $Text = Get-TabHeaderText -Name $Name -Facility $Facility
    if ($IsDirty) {
        $HeaderBlock.Text = "• $Text"
        $HeaderBlock.FontStyle = "Italic"
        $HeaderBlock.Foreground = "#EA580C"
    }
    else {
        $HeaderBlock.Text = $Text
        $HeaderBlock.FontStyle = "Normal"
        $Theme = Get-ThemePalette
        $HeaderBlock.Foreground = $Theme.TextHex
    }
}

function Initialize-EditorWindow {
    $Theme = Get-ThemePalette
    $Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Jelenléti szerkesztő"
        Width="900" Height="650"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="14"
        Background="$($Theme.BackgroundHex)">
    
    <Window.Resources>
        $(Get-SharedButtonStylesXaml)

        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{StaticResource SurfaceBrush}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ListBox">
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBox">
                        <Border BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="4" 
                                Background="{TemplateBinding Background}">
                            <ScrollViewer Focusable="false">
                                <ItemsPresenter/>
                            </ScrollViewer>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border Name="Border" Padding="{TemplateBinding Padding}" Background="Transparent" CornerRadius="3" Margin="4,2">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource HoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource PrimaryBrush}"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="Border" Padding="15,10" BorderThickness="0,0,0,2" BorderBrush="Transparent" Background="Transparent" Margin="0,0,10,0" Cursor="Hand">
                            <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
                                <Setter TargetName="ContentSite" Property="TextElement.Foreground" Value="{StaticResource PrimaryBrush}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource HoverBrush}"/>
                                <Setter TargetName="Border" Property="CornerRadius" Value="4,4,0,0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <DockPanel>
        <Border DockPanel.Dock="Top" Background="{StaticResource SurfaceBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,0,0,1" Padding="20,15">
            <DockPanel>
                <TextBlock Text="Jelenlétik és Igények" FontSize="20" FontWeight="Bold" Foreground="{StaticResource TextBrush}" VerticalAlignment="Center" DockPanel.Dock="Left"/>
                
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" DockPanel.Dock="Right">
                    <Button Name="AddDepartmentButton" Content="+ Új részleg" Width="130" Margin="0,0,10,0"/>
                    <Button Name="ReloadButton" Content="Módosítások elvetése" Width="180" Margin="0,0,10,0"/>
                    <Button Name="SaveButton" Content="Mentés" Width="120" Style="{StaticResource PrimaryButton}"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <Border DockPanel.Dock="Bottom" Background="{StaticResource SurfaceBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,1,0,0" Padding="20,5">
            <TextBlock Name="Status" Text="Készen áll" FontSize="12" Foreground="{StaticResource MutedBrush}"/>
        </Border>

        <Border Padding="20">
            <TabControl Name="TabControl"/>
        </Border>
    </DockPanel>
</Window>
"@

    $Reader = New-Object System.Xml.XmlNodeReader ([xml]$Xaml)
    $global:Window = [Windows.Markup.XamlReader]::Load($Reader)

    $global:SaveButton = $global:Window.FindName("SaveButton")
    $global:ReloadButton = $global:Window.FindName("ReloadButton")
    $global:AddDepartmentButton = $global:Window.FindName("AddDepartmentButton")
    $global:TabControl = $global:Window.FindName("TabControl")
    $global:Status = $global:Window.FindName("Status")

    $global:SaveButton.Add_Click({ Save-All })
    $global:ReloadButton.Add_Click({ Load-All })
    $global:AddDepartmentButton.Add_Click({ Add-Department })

    return $global:Window
}
