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

[System.Reflection.Assembly]::LoadFrom("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll")
Add-Type -AssemblyName PresentationFramework

function Show-LaneMessage {
    param (
        [string]$Title,
        [string]$Message
    )
    [System.Windows.Forms.MessageBox]::Show(
        $(New-Object -TypeName System.Windows.Forms.Form -Property @{ TopMost = $true }),
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null
}

$messages = @{
    "M15_Free" = @{
        Title = "Versenymedence"
        Text  = "Az 50 méteres medence felszabadult"
    }
     "M15_Occ" = @{
        Title = "Versenymedence"
        Text  = "Az 50 méteres medence foglalt"
    }
    "M16_Free" = @{
        Title = "Tanmedence"
        Text  = "A tanmedence felszabadult"
    }
    "M16_Occ" = @{
       Title = "Tanmedence"
        Text  = "A tanmedence foglalt"
    }
    "M17_Free" = @{
        Title = "Bemelegítő medence"
        Text  = "Az 25 méteres medence felszabadult"
    }
    "M17_Occ" = @{
        Title = "Tanmedence""Bemelegítő medence"
        Text  = "Az 25 méteres medence foglalt"
    }
    "M09_Free" = @{
        Title = "Kalandmedence"
        Text  = "A Kalandmedence felszabadult"
    }
     "M09_Occ" = @{
        Title = "Kalandmedence"
        Text  = "A Kalandmedence foglalt"
    }
}

if ($messages.ContainsKey($MessageID)) {
    $msg = $messages[$MessageID]
    Show-LaneMessage -Title $msg.Title -Message $msg.Text
}