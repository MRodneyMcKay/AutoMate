param (
    [string]$ReminderID
)
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

$messages = @{
    "TIG" = @{
        Title = "Teljesítés igazolás"
        Text  = "Ellenőrzd a Teljesítés Igazolásokon a létszámokat!"
    }
    "Csúszda" = @{
        Title = "Csúszda"
        Text  = "A csúszda kinyitott"
    }
    "Szauna" = @{
        Title = "Szauna"
        Text  = "A szauna kinyitott"
    }
}

if ($messages.ContainsKey($ReminderID)) {
    $msg = $messages[$ReminderID]
    #Show-Message -Title $msg.Title -Message $msg.Text
    Write-Log -Message "Show message box for reminder:" -Level "INFO" -ShowMessageBox -MsgBoxTitle $msg.Title -MsgBoxMessage $msg.Text
}