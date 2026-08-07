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

function Get-ThemeResources {
    [CmdletBinding()]
    param()

    $AccentColor = [System.Windows.SystemColors]::AccentColor

    try {
        $ThemeKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        $AppsUseLightTheme = Get-ItemPropertyValue -Path $ThemeKey -Name 'AppsUseLightTheme' -ErrorAction Stop
        $IsDark = ([int]$AppsUseLightTheme -eq 0)
    }
    catch {
        $IsDark = $false
    }

    if ($IsDark) {
        $accentActiveBrush = [System.Windows.SystemColors]::AccentColorLight3Brush
        $controlBgColor = [System.Windows.Media.Color]::FromRgb(0x2D, 0x2D, 0x30)
        $BackgroundColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(18, 18, 18)) -Ratio 0.10
        $SurfaceColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(30, 30, 30)) -Ratio 0.18
        $BorderColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(90, 90, 90)) -Ratio 0.24
        $TextColor = [System.Windows.Media.Color]::FromRgb(248, 250, 252)
        $MutedColor = [System.Windows.Media.Color]::FromRgb(203, 213, 225)
        $HoverColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(48, 48, 48)) -Ratio 0.28
        $PressedColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(40, 40, 40)) -Ratio 0.32
        $SelectionColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(60, 60, 60)) -Ratio 0.30
    }
    else {
        $accentActiveBrush = [System.Windows.SystemColors]::AccentColorDark3Brush
        $controlBgColor = [System.Windows.Media.Color]::FromRgb(255, 255, 255)
        $BackgroundColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(248, 250, 252)) -Ratio 0.08
        $SurfaceColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(255, 255, 255)) -Ratio 0.04
        $BorderColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(209, 213, 219)) -Ratio 0.12
        $TextColor = [System.Windows.Media.Color]::FromRgb(17, 24, 39)
        $MutedColor = [System.Windows.Media.Color]::FromRgb(107, 114, 128)
        $HoverColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(249, 250, 251)) -Ratio 0.10
        $PressedColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(229, 231, 235)) -Ratio 0.14
        $SelectionColor = Merge-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(239, 244, 255)) -Ratio 0.20
    }

    $controlHoverColor = Lighten $controlBgColor 0.10
    $controlPressedColor = Darken  $controlBgColor 0.12
    $accentHoverColor = Lighten $AccentColor    0.12
    $accentPressedColor = Darken  $AccentColor    0.18
    $AccentForegroundColor = Get-AccessibleAccentForeground -AccentColor $AccentColor
    $DropHighlightColor = Merge-Colors -BaseColor $AccentColor -OverlayColor $SurfaceColor -Ratio 0.35
    $rd = [System.Windows.ResourceDictionary]::new()

    function Add-Brush {
        param(
            [System.Windows.ResourceDictionary]$Dictionary,
            [string]$Key,
            [System.Windows.Media.Color]$Color
        )
        $Dictionary[$Key] = [System.Windows.Media.SolidColorBrush]::new($Color)
    }

    function Add-String {
        param($Dictionary, $Key, $Value)
        $Dictionary[$Key] = $Value
    }

    function Add-Bool {
        param($Dictionary, $Key, [bool]$Value)
        $Dictionary[$Key] = $Value
    }

    # Brushes
    Add-Brush $rd AccentBrush            $AccentColor
    Add-Brush $rd AccentForegroundBrush  $AccentForegroundColor
    Add-Brush $rd BackgroundBrush        $BackgroundColor
    Add-Brush $rd SurfaceBrush           $SurfaceColor
    Add-Brush $rd BorderBrush            $BorderColor
    Add-Brush $rd TextBrush              $TextColor
    Add-Brush $rd MutedBrush             $MutedColor
    Add-Brush $rd HoverBrush             $HoverColor
    Add-Brush $rd PressedBrush           $PressedColor
    Add-Brush $rd SelectionBrush         $SelectionColor
    Add-Brush $rd ControlHoverBrush      $controlHoverColor
    Add-Brush $rd ControlPressedBrush    $controlPressedColor
    Add-Brush $rd AccentHoverBrush       $accentHoverColor
    Add-Brush $rd AccentPressedBrush     $accentPressedColor
    Add-Brush $rd AccentActiveBrush      $accentActiveBrush.Color
    Add-Brush $rd DropHighlightBrush     $DropHighlightColor

    Add-String $rd AccentHex             (Convert-ColorToHex $AccentColor)
    Add-String $rd AccentForegroundHex   (Convert-ColorToHex $AccentForegroundColor)
    Add-String $rd BackgroundHex         (Convert-ColorToHex $BackgroundColor)
    Add-String $rd SurfaceHex            (Convert-ColorToHex $SurfaceColor)
    Add-String $rd BorderHex             (Convert-ColorToHex $BorderColor)
    Add-String $rd TextHex               (Convert-ColorToHex $TextColor)
    Add-String $rd MutedHex              (Convert-ColorToHex $MutedColor)
    Add-String $rd HoverHex              (Convert-ColorToHex $HoverColor)
    Add-String $rd PressedHex            (Convert-ColorToHex $PressedColor)
    Add-String $rd SelectionHex          (Convert-ColorToHex $SelectionColor)
    Add-String $rd ControlHoverHex       (Convert-ColorToHex $controlHoverColor)
    Add-String $rd ControlPressedHex     (Convert-ColorToHex $controlPressedColor)
    Add-String $rd AccentHoverHex        (Convert-ColorToHex $accentHoverColor)
    Add-String $rd AccentPressedHex      (Convert-ColorToHex $accentPressedColor)
    Add-String $rd AccentActiveHex       (Convert-ColorToHex $accentActiveBrush.Color)

    # Other values
    Add-Bool $rd IsDark $IsDark

    return $rd
}

function Get-ThemeStyle {
    return  [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml](Get-Content (Join-Path $PSScriptRoot "Styles/ThemeStyles.xaml") -Raw))))
}