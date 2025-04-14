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
. $PSScriptRoot\Private\nationalityCodes.ps1

#import public functions
. $PSScriptRoot\Public\nationalities.ps1

Export-ModuleMember -Function Get-GroupedNationalityData, Get-NationalityStats

Write-Log -Message "Module loaded: VisitorStatistic"