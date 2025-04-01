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

Import-Module "$PSScriptRoot\MonthlyCommutingAllowancePrint.psm1"
Import-Module "$PSScriptRoot\monthlyPrintAttendceSheetCollegues.psm1"
Import-Module "$PSScriptRoot\monthlyPrintScheduleRequest.psm1"
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\Log.psm1')

$path = Get-SheetPath

Write-Log -Message "Printing scedule requests for the front office"
Print-RequestFrontOffice
Write-Log -Message "Printing attendance sheets for the front office"
Print-AttandanceSheetFrontOffice -OpenFile $path

Write-Log -Message "Printing scedule requests for the staff"
Print-RequestUszomester
Write-Log -Message "Printing commuting allowance"
Print-CommutingAllowance
Write-Log -Message "Printing attendance sheets for the staff"
Print-AttandanceSheetUszomester -OpenFile $path

Write-Log -Message "Printing attendance sheets for the genitors"
Print-AttandanceSheetKarbantarto -OpenFile $path
Write-Log -Message "Printing attendance sheets for the pool technicians"
Print-AttandanceSheetGepesz -OpenFile $path