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

Import-Module "$PSScriptRoot\wallpaperchanger.psm1"
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\Log.psm1')

# Initialize common personalization settings
function Initialize-Settings {
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name AutoColorization -Value 1 -Type Dword -Force
    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\DWM" -Name ColorPrevalence -Value 1 -Type Dword -Force
    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name EnableTransparency -Value 1 -Type Dword -Force
}

# Set personalization values
function Set-Personalization {
    param (
        [string]$AppsUseLightTheme,
        [string]$SystemUsesLightTheme,
        [string]$ColorPrevalence,
        [string]$PicturePath,
        [string]$Nickname
    )
    Initialize-Settings
    $registryPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    New-ItemProperty -Path $registryPath -Name AppsUseLightTheme -Value $AppsUseLightTheme -Type Dword -Force
    New-ItemProperty -Path $registryPath -Name SystemUsesLightTheme -Value $SystemUsesLightTheme -Type Dword -Force
    New-ItemProperty -Path $registryPath -Name ColorPrevalence -Value $ColorPrevalence -Type Dword -Force
    Set-DesktopWallpaper -PicturePath $PicturePath -Style Fill
    Write-Log -Message "Theme applied for $Nickname"
}

# Functions to change themes and wallpapers
function Set-Default {
    Set-Personalization -AppsUseLightTheme 1 -SystemUsesLightTheme 1 -ColorPrevalence 0 -PicturePath "C:\Windows\Web\Wallpaper\Windows\img0.jpg" -Nickname "Default"
}

function Set-Zsolti {
    Set-Personalization -AppsUseLightTheme 0 -SystemUsesLightTheme 0 -ColorPrevalence 1 -PicturePath "C:\Windows\Web\Wallpaper\Windows\img1.jpg" -Nickname "Zsolti"
}

function Set-Niki {
    Set-Personalization -AppsUseLightTheme 0 -SystemUsesLightTheme 0 -ColorPrevalence 1 -PicturePath "C:\Windows\Web\Wallpaper\Windows\img3.jpg" -Nickname "Niki"
}