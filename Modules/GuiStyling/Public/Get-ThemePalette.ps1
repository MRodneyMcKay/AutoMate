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


function Convert-ColorToHex {
    [CmdletBinding()]
    param(
        [System.Windows.Media.Color]$Color
    )

    return "#{0:X2}{1:X2}{2:X2}" -f $Color.R, $Color.G, $Color.B
}

function Blend-Colors {
    [CmdletBinding()]
    param(
        [System.Windows.Media.Color]$BaseColor,
        [System.Windows.Media.Color]$OverlayColor,
        [double]$Ratio
    )

    $Ratio = [Math]::Min(1.0, [Math]::Max(0.0, $Ratio))
    $R = [int][Math]::Round(($BaseColor.R * $Ratio) + ($OverlayColor.R * (1 - $Ratio)))
    $G = [int][Math]::Round(($BaseColor.G * $Ratio) + ($OverlayColor.G * (1 - $Ratio)))
    $B = [int][Math]::Round(($BaseColor.B * $Ratio) + ($OverlayColor.B * (1 - $Ratio)))

    return [System.Windows.Media.Color]::FromRgb([byte]$R, [byte]$G, [byte]$B)
}

function Get-ThemePalette {
    [CmdletBinding()]
    param()

    $AccentColor = [System.Windows.Media.Color]::FromArgb(
        255,
        [byte]([System.Windows.SystemParameters]::WindowGlassColor.R),
        [byte]([System.Windows.SystemParameters]::WindowGlassColor.G),
        [byte]([System.Windows.SystemParameters]::WindowGlassColor.B)
    )

    $IsDark = $false
    try {
        $ThemeKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        $AppsUseLightTheme = Get-ItemPropertyValue -Path $ThemeKey -Name 'AppsUseLightTheme' -ErrorAction Stop
        $IsDark = [int]$AppsUseLightTheme -eq 0
    }
    catch {
        $IsDark = $false
    }

    if ($IsDark) {
        $BackgroundColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(18, 18, 18)) -Ratio 0.10
        $SurfaceColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(30, 30, 30)) -Ratio 0.18
        $BorderColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(90, 90, 90)) -Ratio 0.24
        $TextColor = [System.Windows.Media.Color]::FromRgb(248, 250, 252)
        $MutedColor = [System.Windows.Media.Color]::FromRgb(203, 213, 225)
        $HoverColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(48, 48, 48)) -Ratio 0.28
        $PressedColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(40, 40, 40)) -Ratio 0.32
        $SelectionColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(60, 60, 60)) -Ratio 0.30
    }
    else {
        $BackgroundColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(248, 250, 252)) -Ratio 0.08
        $SurfaceColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(255, 255, 255)) -Ratio 0.04
        $BorderColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(209, 213, 219)) -Ratio 0.12
        $TextColor = [System.Windows.Media.Color]::FromRgb(17, 24, 39)
        $MutedColor = [System.Windows.Media.Color]::FromRgb(107, 114, 128)
        $HoverColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(249, 250, 251)) -Ratio 0.10
        $PressedColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(229, 231, 235)) -Ratio 0.14
        $SelectionColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(239, 244, 255)) -Ratio 0.20
    }

    $PrimaryHoverColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(255, 255, 255)) -Ratio 0.14
    $PrimaryPressedColor = Blend-Colors -BaseColor $AccentColor -OverlayColor ([System.Windows.Media.Color]::FromRgb(0, 0, 0)) -Ratio 0.16

    return [ordered]@{
        AccentHex = Convert-ColorToHex -Color $AccentColor
        AccentHoverHex = Convert-ColorToHex -Color $PrimaryHoverColor
        AccentPressedHex = Convert-ColorToHex -Color $PrimaryPressedColor
        BackgroundHex = Convert-ColorToHex -Color $BackgroundColor
        SurfaceHex = Convert-ColorToHex -Color $SurfaceColor
        BorderHex = Convert-ColorToHex -Color $BorderColor
        TextHex = Convert-ColorToHex -Color $TextColor
        MutedHex = Convert-ColorToHex -Color $MutedColor
        HoverHex = Convert-ColorToHex -Color $HoverColor
        PressedHex = Convert-ColorToHex -Color $PressedColor
        SelectionHex = Convert-ColorToHex -Color $SelectionColor
        IsDark = $IsDark
    }
}
