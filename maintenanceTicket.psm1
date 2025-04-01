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

function New-Ticket {
    
$shiftManager = (Get-ShiftManager)[(Get-Date).TimeOfDay -lt (New-TimeSpan -Hours 12 -Minutes 55) ? 0 : 1].Name


$today = Get-Date
$Excel = New-Object -ComObject Excel.Application
$Excel.visible=$false

Invoke-Item "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Email sablonok\HIBA.oft"

$Workbook = $Excel.Workbooks.Open("C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\HIBA\-=SABLON=-\HIBA jelentés.xlsx")
$workSheet = $Workbook.Sheets.Item("Munka1")
$workSheet.Range("A6").value()= Get-Date $today -Format "yyyy.MM.dd. HH:mm"
$workSheet.Range("A23").value()= $shiftManager
$Excel.visible=$true
}