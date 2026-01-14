param (
    [string]$MessageID
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
    "M02_Free" = @{
        Title = "Panoráma medence"
        Text  = "A panoráma medence felszabadult"
    }
    "M09_Free" = @{
        Title = "Kalandmedence"
        Text  = "A Kalandmedence felszabadult"
    }
    "M15_Free" = @{
        Title = "Versenymedence"
        Text  = "Az 50 méteres úszómedence felszabadult"
    }
    "M16_Free" = @{
        Title = "Tanmedence"
        Text  = "A tanmedence felszabadult"
    }
    "M17_Free" = @{
        Title = "Bemelegítőmedence"
        Text  = "A 25 méteres úszómedence felszabadult"
    }
    "M02_Occ" = @{
        Title = "Panoráma medence"
        Text  = "A panoráma medence foglalt"
    }  
    "M09_Occ" = @{
        Title = "Kalandmedence"
        Text  = "A Kalandmedence foglalt"
    }    
    "M15_Occ" = @{
        Title = "Versenymedence"
        Text  = "Az 50 méteres úszómedence foglalt"
    }    
    "M16_Occ" = @{
        Title = "Tanmedence"
        Text  = "A tanmedence foglalt"
    }    
    "M17_Occ" = @{
        Title = "Bemelegítőmedence"
        Text  = "A 25 méteres úszómedence foglalt"
    }
}

if ($messages.ContainsKey($MessageID)) {
    $msg = $messages[$MessageID]
    Write-Log -Message "Show message box for pool schedule:" -Level "INFO" -ShowMessageBox -MsgBoxTitle $msg.Title -MsgBoxMessage $msg.Text
}