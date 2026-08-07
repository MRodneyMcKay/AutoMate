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

$PrintActions = @{
    "PrintRequestFrontOffice"     = { Write-Log 'Printing schedule requests for the front office'; Print-RequestFrontOffice }
    "PrintAttendanceFrontOffice"  = { param($SheetPath); Write-Log 'Printing attendance sheets for the front office'; Print-AttandanceSheetFrontOffice -OpenFile $SheetPath }

    "PrintRequestUszomester"      = { Write-Log 'Printing schedule requests for the staff'; Print-RequestUszomester }
    "PrintAttendanceUszomester"   = { param($SheetPath); Write-Log 'Printing attendance sheets for the staff'; Print-AttandanceSheetUszomester -OpenFile $SheetPath }

    "PrintCommutingAllowance"     = { Write-Log 'Printing commuting allowance'; Print-CommutingAllowance }

    "PrintRequestDombBeach"       = { Write-Log 'Printing schedule requests for Domb Beach'; Print-RequestDombBeach }

    "PrintAttendanceKarbantarto"  = { param($SheetPath); Write-Log 'Printing attendance sheets for the genitors'; Print-AttandanceSheetKarbantarto -OpenFile $SheetPath }

    "PrintAttendanceGepesz"       = { param($SheetPath); Write-Log 'Printing attendance sheets for the pool technicians'; Print-AttandanceSheetGepesz -OpenFile $SheetPath }

    "PrintRequestGyogyaszat"      = { Write-Log 'Printing schedule requests for the medical department'; Print-RequestGyogyaszat }
    "PrintAttendanceGyogyaszat"   = { param($SheetPath); Write-Log 'Printing attendance sheets for the medical department'; Print-AttandanceSheetGyogyaszat -OpenFile $SheetPath }

    "PrintAttendanceOnkormanyzat" = { param($SheetPath); Write-Log 'Printing attendance sheets for the municipality'; Print-AttandanceSheetOnkormanyzat -OpenFile $SheetPath }
}
