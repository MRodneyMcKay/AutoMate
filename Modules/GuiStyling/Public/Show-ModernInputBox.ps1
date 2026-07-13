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


function Show-ModernInputBox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )

    $Theme = Get-ThemePalette
    $DialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Width="450" SizeToContent="Height"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="14" Background="$($Theme.BackgroundHex)">
    <Window.Resources>
        $(Get-SharedButtonStylesXaml)
        
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Border" Background="{StaticResource SurfaceBrush}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Border Padding="25">
        <StackPanel>
            <TextBlock Text="$Prompt" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,15" TextWrapping="Wrap"/>
            <TextBox Name="InputBox" Margin="0,0,0,25"/>
            
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="CancelBtn" Content="Mégse" Width="90" Margin="0,0,10,0"/>
                <Button Name="OkBtn" Content="Rendben" Width="100" Style="{StaticResource PrimaryButton}"/>
            </StackPanel>
        </StackPanel>
    </Border>
</Window>
"@

    $Reader = New-Object System.Xml.XmlNodeReader ([xml]$DialogXaml)
    $Dialog = [Windows.Markup.XamlReader]::Load($Reader)

    $InputBox = $Dialog.FindName('InputBox')
    $OkBtn = $Dialog.FindName('OkBtn')
    $CancelBtn = $Dialog.FindName('CancelBtn')

    $Dialog.Add_Loaded({ $InputBox.Focus() })

    $OkBtn.Add_Click({ $Dialog.DialogResult = $true })
    $CancelBtn.Add_Click({ $Dialog.DialogResult = $false })

    $InputBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq 'Enter') {
            $Dialog.DialogResult = $true
        }
        elseif ($e.Key -eq 'Escape') {
            $Dialog.DialogResult = $false
        }
    })

    if ($Dialog.ShowDialog() -eq $true) {
        return $InputBox.Text.Trim()
    }

    return ''
}
