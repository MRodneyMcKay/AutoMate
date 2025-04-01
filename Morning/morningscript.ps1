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

$registryPath = "HKCU:\Software\Script"

Import-Module "$PSScriptRoot\createTemplate.psm1"
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\RosterInformation.psm1')
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\Log.psm1')
Import-Module "$PSScriptRoot\createDir.psm1"
Import-Module "$PSScriptRoot\createWorkingHoursExcel.psm1"
Import-Module "$PSScriptRoot\openEmails.psm1"

create-Directories

if ($(Get-ItemPropertyValue -Path $registryPath -Name EmailLastShown) -ne (Get-Date).Day) {
    Write-Log -Message "Emailek megnyitása"
    Open-Emails
    Set-ItemProperty -Path $registryPath -Name EmailLastShown -value (Get-Date).Day
} else {
    Write-Log -Message "Emailek már megnyitva voltak"
}
if ($(Get-ItemPropertyValue -Path $registryPath -Name YesterdaysWorkingHours) -ne (Get-Date).Day) {
    Write-Log -Message "Diákelszámolás elkészítése"
    $roster = Get-Receptionists
    $approved = @(1, 112, "1/9")
    open-StudentWorkReportFurdo -fillCompletely ((($roster | Where-Object { $_.Name -eq 'Raduska Zsolt' }).Shift -in $approved -or ($roster | Where-Object { $_.Name -eq 'Konfár Nikolett' }).Shift -in $approved))
    Set-ItemProperty -Path $registryPath -Name YesterdaysWorkingHours -value (Get-Date).Day
} else {
    Write-Log -Message "Diákelszámolás már elkészült"
}

Start-Process msedge

Set-ItemProperty -Path $registryPath -Name LastRun -value (Get-Date).Day