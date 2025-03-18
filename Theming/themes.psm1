Import-Module "$PSScriptRoot\wallpaperchanger.psm1"

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
        [string]$PicturePath
    )
    Initialize-Settings
    $registryPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    New-ItemProperty -Path $registryPath -Name AppsUseLightTheme -Value $AppsUseLightTheme -Type Dword -Force
    New-ItemProperty -Path $registryPath -Name SystemUsesLightTheme -Value $SystemUsesLightTheme -Type Dword -Force
    New-ItemProperty -Path $registryPath -Name ColorPrevalence -Value $ColorPrevalence -Type Dword -Force
    Set-DesktopWallpaper -PicturePath $PicturePath -Style Fill
}

# Functions to change themes and wallpapers
function Set-Default {
    Set-Personalization -AppsUseLightTheme 1 -SystemUsesLightTheme 1 -ColorPrevalence 0 -PicturePath "C:\Windows\Web\Wallpaper\Windows\img0.jpg"
}

function Set-Zsolti {
    Set-Personalization -AppsUseLightTheme 0 -SystemUsesLightTheme 0 -ColorPrevalence 1 -PicturePath "C:\Windows\Web\Wallpaper\Windows\img1.jpg"
}

function Set-Niki {
    Set-Personalization -AppsUseLightTheme 0 -SystemUsesLightTheme 0 -ColorPrevalence 1 -PicturePath "C:\Windows\Web\Wallpaper\Windows\img3.jpg"
}