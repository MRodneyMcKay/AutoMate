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

Import-Module $PSScriptRoot\RosterInformation.psm1

[System.Reflection.Assembly]::LoadFrom("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll")
Add-Type -AssemblyName PresentationFramework

$shiftManager =  (Get-ShiftManager)[(Get-Date).TimeOfDay -lt (New-TimeSpan -Hours 12 -Minutes 55) ? 0 : 1]

if ($null -eq $shiftManager.Name) {
    [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "Nem tudom ki a $(($shiftManager.Shift -eq "1~*") ? "délelöttös" : "délutános") műszakvezető", "Vajon ki a müszi?", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Error)
}
else {
    [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "A $(($shiftManager.Shift -eq "1*") ? "délelöttös" : "délutános") műszakvezető: $($shiftManager.Name)", "Vajon ki a müszi?", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Information)
}