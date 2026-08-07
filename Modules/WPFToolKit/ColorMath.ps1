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

function Merge-Colors {
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

function Get-RelativeLuminance {
    [CmdletBinding()]
    param(
        [System.Windows.Media.Color]$Color
    )

    $normalize = {
        param([double]$Channel)
        $Channel = $Channel / 255.0
        if ($Channel -le 0.03928) {
            return $Channel / 12.92
        }
        return [Math]::Pow((($Channel + 0.055) / 1.055), 2.4)
    }

    $R = & $normalize -Channel $Color.R
    $G = & $normalize -Channel $Color.G
    $B = & $normalize -Channel $Color.B

    return 0.2126 * $R + 0.7152 * $G + 0.0722 * $B
}

function Get-ContrastRatio {
    [CmdletBinding()]
    param(
        [System.Windows.Media.Color]$Foreground,
        [System.Windows.Media.Color]$Background
    )

    $L1 = Get-RelativeLuminance -Color $Foreground
    $L2 = Get-RelativeLuminance -Color $Background

    if ($L1 -lt $L2) {
        return ($L1 + 0.05) / ($L2 + 0.05)
    }

    return ($L2 + 0.05) / ($L1 + 0.05)
}

function Resolve-AccessibleAccentColor {
    [CmdletBinding()]
    param(
        [System.Windows.Media.Color]$AccentColor,
        [System.Windows.Media.Color]$BackgroundColor,
        [bool]$IsDark
    )

    if (-not $IsDark) {
        return $AccentColor
    }

    $overlayColor = [System.Windows.Media.Color]::FromRgb(255, 255, 255)
    $candidate = $AccentColor

    for ($ratio = 0.0; $ratio -le 1.0; $ratio += 0.05) {
        $candidate = Merge-Colors -BaseColor $AccentColor -OverlayColor $overlayColor -Ratio $ratio
        if ((Get-ContrastRatio -Foreground $candidate -Background $BackgroundColor) -ge 4.5) {
            return $candidate
        }
    }

    return $candidate
}

function Lighten {
    param([System.Windows.Media.Color]$c, [double]$factor)
    [System.Windows.Media.Color]::FromRgb(
        [byte]([Math]::Min(255, $c.R + 255 * $factor)),
        [byte]([Math]::Min(255, $c.G + 255 * $factor)),
        [byte]([Math]::Min(255, $c.B + 255 * $factor))
    )
}

function Darken {
    param([System.Windows.Media.Color]$c, [double]$factor)
    [System.Windows.Media.Color]::FromRgb(
        [byte]([Math]::Max(0, $c.R - 255 * $factor)),
        [byte]([Math]::Max(0, $c.G - 255 * $factor)),
        [byte]([Math]::Max(0, $c.B - 255 * $factor))
    )
}

function ConvertTo-HSL {
    param([System.Windows.Media.Color]$Color)

    $r = $Color.R / 255.0
    $g = $Color.G / 255.0
    $b = $Color.B / 255.0

    $max = [Math]::Max($r, [Math]::Max($g, $b))
    $min = [Math]::Min($r, [Math]::Min($g, $b))
    $delta = $max - $min

    # Lightness
    $l = ($max + $min) / 2.0

    if ($delta -eq 0) {
        return [pscustomobject]@{ H = 0; S = 0; L = $l }
    }

    # Saturation
    if ($l -lt 0.5) {
        $s = $delta / ($max + $min)
    }
    else {
        $s = $delta / (2.0 - $max - $min)
    }

    # Hue
    if ($max -eq $r) {
        $h = (($g - $b) / $delta) % 6
    }
    elseif ($max -eq $g) {
        $h = (($b - $r) / $delta) + 2
    }
    else {
        $h = (($r - $g) / $delta) + 4
    }

    $h = $h * 60
    if ($h -lt 0) { $h += 360 }

    return [pscustomobject]@{
        H = $h
        S = $s
        L = $l
    }
}

function ConvertFrom-HSL {
    param(
        [double]$H,
        [double]$S,
        [double]$L
    )

    $C = (1 - [Math]::Abs(2 * $L - 1)) * $S
    $X = $C * (1 - [Math]::Abs((($H / 60) % 2) - 1))
    $m = $L - $C / 2

    switch ($H) {
        { $_ -lt 60 } { $r = $C; $g = $X; $b = 0 }
        { $_ -lt 120 } { $r = $X; $g = $C; $b = 0 }
        { $_ -lt 180 } { $r = 0; $g = $C; $b = $X }
        { $_ -lt 240 } { $r = 0; $g = $X; $b = $C }
        { $_ -lt 300 } { $r = $X; $g = 0; $b = $C }
        default { $r = $C; $g = 0; $b = $X }
    }

    return [System.Windows.Media.Color]::FromRgb(
        [byte]([Math]::Round(($r + $m) * 255)),
        [byte]([Math]::Round(($g + $m) * 255)),
        [byte]([Math]::Round(($b + $m) * 255))
    )
}

function Get-AccessibleAccentForeground {
    param([System.Windows.Media.Color]$AccentColor)

    $hsl = ConvertTo-HSL $AccentColor

    # Komplementer hue
    $h = ($hsl.H + 180) % 360

    # Kezdő paraméterek
    $s = 0.35
    $l = 0.75

    $candidate = ConvertFrom-HSL -H $h -S $s -L $l
    $contrast = Get-ContrastRatio -Foreground $candidate -Background $AccentColor

    # Ha nem elég kontrasztos → finomhangoljuk a lightness-t
    if ($contrast -lt 4.5) {
        for ($i = 0.75; $i -ge 0.1; $i -= 0.05) {
            $candidate = ConvertFrom-HSL -H $h -S $s -L $i
            $contrast = Get-ContrastRatio -Foreground $candidate -Background $AccentColor
            if ($contrast -ge 4.5) {
                return $candidate
            }
        }
    }

    return $candidate
}