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


function Get-SharedButtonStylesXaml {
    [CmdletBinding()]
    param()

    return Get-WpfSharedStylesXaml -IncludeButtonStyles
}

function Get-WpfSharedStylesXaml {
    [CmdletBinding()]
    param(
        [switch]$IncludeButtonStyles,
        [switch]$IncludeTextBoxStyles,
        [switch]$IncludeListBoxStyles,
        [switch]$IncludeTabStyles
    )

    $Theme = Get-ThemePalette
    $parts = [System.Collections.Generic.List[string]]::new()

    $parts.Add(@"
        <SolidColorBrush x:Key="PrimaryBrush" Color="$($Theme.AccentHex)"/>
        <SolidColorBrush x:Key="PrimaryHoverBrush" Color="$($Theme.AccentHoverHex)"/>
        <SolidColorBrush x:Key="PrimaryPressedBrush" Color="$($Theme.AccentPressedHex)"/>
        <SolidColorBrush x:Key="BackgroundBrush" Color="$($Theme.BackgroundHex)"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="$($Theme.SurfaceHex)"/>
        <SolidColorBrush x:Key="BorderBrush" Color="$($Theme.BorderHex)"/>
        <SolidColorBrush x:Key="TextBrush" Color="$($Theme.TextHex)"/>
        <SolidColorBrush x:Key="MutedBrush" Color="$($Theme.MutedHex)"/>
        <SolidColorBrush x:Key="HoverBrush" Color="$($Theme.HoverHex)"/>
        <SolidColorBrush x:Key="PressedBrush" Color="$($Theme.PressedHex)"/>
        <SolidColorBrush x:Key="SelectionBrush" Color="$($Theme.SelectionHex)"/>
"@)

    if ($IncludeButtonStyles -or -not ($IncludeTextBoxStyles -or $IncludeListBoxStyles -or $IncludeTabStyles)) {
        $parts.Add(@"
        <Style TargetType="Button">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource HoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{StaticResource PressedBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{StaticResource PrimaryBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource PrimaryHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{StaticResource PrimaryPressedBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
"@)
    }

    if ($IncludeTextBoxStyles) {
        $parts.Add(@"
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{StaticResource SurfaceBrush}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
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
"@)
    }

    if ($IncludeListBoxStyles) {
        $parts.Add(@"
        <Style TargetType="ListBox">
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBox">
                        <Border BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4" Background="{TemplateBinding Background}">
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
"@)
    }

    if ($IncludeTabStyles) {
        $parts.Add(@"
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
"@)
    }

    return ($parts -join [Environment]::NewLine)
}
