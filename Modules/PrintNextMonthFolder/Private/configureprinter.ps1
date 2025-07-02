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

function Configure-Printer {
    param (
        [string]$PrinterName = (Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }).Name,
        [string]$DuplexingMode
    )
    try {
        Set-PrintConfiguration -PrinterName $PrinterName -DuplexingMode $DuplexingMode | Out-Null
        Write-Log -Message "Configure printer $PrinterName with duplexing mode $DuplexingMode." -Level INFO
    } catch {
        Write-Log -Message "Failed to configure printer $PrinterName with duplexing mode $DuplexingMode. Error: $_" -Level Error
    }
}