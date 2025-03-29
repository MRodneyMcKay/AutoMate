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

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile = "C:\Logs\log.txt",
        [string]$timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    )
    $entry = @(
        "Date: $(Get-Date -Format 'yyyy-MM-dd')",
        "Time: $(Get-Date -Format 'HH:mm:ss')",
        "Traceback: $((Get-PSCallStack)[1].ScriptName):$((Get-PSCallStack)[1].FunctionName)",
        "Message: $Message"
    )

   # Append log entry to the file
   Add-Content -Path $LogFile -Value $entry
}