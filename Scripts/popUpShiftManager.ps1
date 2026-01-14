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
$resolvedModulePath = (Resolve-Path -Path $ModulePath).Path

Import-Module (Join-Path -Path $resolvedModulePath -ChildPath 'LoggingSystem\LoggingSystem.psd1')
Import-Module (Join-Path -Path $resolvedModulePath -ChildPath 'RosterInformation\RosterInformation.psd1')

try {
    $shiftManager = Get-ShiftManager
}
catch {
    Write-Log `
        -Message "Nem tudom ki most a műszakvezető" `
        -Level ERROR `
        -ShowMessageBox `
        -MsgBoxTitle "Vajon ki a müszi?" `
        -MsgBoxMessage "Nem tudom ki most a műszakvezető."
    exit
}

# Meghatározzuk, hogy délelőttös vagy délutános
$shiftType = ((Get-Date).TimeOfDay -lt (New-TimeSpan -Hours 12 -Minutes 55)) ?
    "délelőttös" :
    "délutános"

Write-Log `
    -Message "A $shiftType műszakvezető: $shiftManager" `
    -Level INFO `
    -ShowMessageBox `
    -MsgBoxTitle "Vajon ki a müszi?" `
    -MsgBoxMessage "A $shiftType műszakvezető: $shiftManager"

