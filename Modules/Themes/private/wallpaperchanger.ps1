<# This file is based on "Set-Wall.ps1" originally created by 5nik7.
Source: https://github.com/5nik7/dots/blob/c888efe778472529dd73414de086843bd7e0dcef/shells/powershell/Scripts/Set-Wall.ps1
Licensed under the MIT License. See the license text below.

Modifications made by MRodneyMcKay, 2025.
Changes include: adding a wrapper function and making the script into a module.

MIT License
Copyright (c) 2024 5nik7

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE. #>

function Set-DesktopWallpaper {
    <#
    .DESCRIPTION
        Sets a desktop background image
    .PARAMETER PicturePath
        Defines the path to the picture to use for background
    .PARAMETER Style
        Defines the style of the wallpaper. Valid values are, Tiled, Centered, Stretched, Fill, Fit, Span
    .EXAMPLE
        Set-DesktopWallpaper -PicturePath "C:\pictures\picture1.jpg" -Style Fill
    .EXAMPLE
        Set-DesktopWallpaper -PicturePath "C:\pictures\picture2.png" -Style Centered
    .NOTES
        Supports jpg, png and bmp files.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$PicturePath,
        [ValidateSet('Tiled', 'Centered', 'Stretched', 'Fill', 'Fit', 'Span')]$Style = 'Fill'
    )


    BEGIN {
        $Definition = @"
[DllImport("user32.dll", EntryPoint = "SystemParametersInfo")]
public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
"@

        Add-Type -MemberDefinition $Definition -Name Win32SystemParametersInfo -Namespace Win32Functions
        $Action_SetDeskWallpaper = [int]20
        $Action_UpdateIniFile = [int]0x01
        $Action_SendWinIniChangeEvent = [int]0x02

        $HT_WallPaperStyle = @{
            'Tiles'     = 0
            'Centered'  = 0
            'Stretched' = 2
            'Fill'      = 10
            'Fit'       = 6
            'Span'      = 22
        }

        $HT_TileWallPaper = @{
            'Tiles'     = 1
            'Centered'  = 0
            'Stretched' = 0
            'Fill'      = 0
            'Fit'       = 0
            'Span'      = 0
        }

    }


    PROCESS {
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name wallpaperstyle -Value $HT_WallPaperStyle[$Style]
        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name tilewallpaper -Value $HT_TileWallPaper[$Style]
        $null = [Win32Functions.Win32SystemParametersInfo]::SystemParametersInfo($Action_SetDeskWallpaper, 0, $PicturePath, ($Action_UpdateIniFile -bor $Action_SendWinIniChangeEvent))
    }
}