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

$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$resolvedModulegPath = (Resolve-Path -Path $ModulePath).Path

Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'LoggingSystem\LoggingSystem.psd1')
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'Themes\themes.psd1')

# Get the value of the environment variable
$theme = [System.Environment]::GetEnvironmentVariable('ApplyTheme', [System.EnvironmentVariableTarget]::User)

# Map the user choice to the function name and invoke the function using the call operator (&)
$functionName = "Set-$theme"
if (Get-Command $functionName -ErrorAction SilentlyContinue) {
    & $functionName
    Write-Log -Message "témaváltás sikeresen kezdeményezve"
} else {
    Write-Log -Message "Nem tudom milyen témát állítsak be: $theme" -Level "ERROR"
}
