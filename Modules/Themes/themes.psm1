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

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\LoggingSystem\LoggingSystem.psd1')

#import provate functions
. $PSScriptRoot\Private\wallpaperchanger.ps1
. $PSScriptRoot\Private\internalhelper.ps1

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

Export-ModuleMember -Function Set-Default, Set-Zsolti, Set-Niki

Write-Log -Message "Module loaded: themes"