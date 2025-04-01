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

[System.Reflection.Assembly]::LoadFrom([System.Environment]::GetEnvironmentVariable("OfficeAssemblies_Outlook", [System.EnvironmentVariableTarget]::User)) 
function Open-OutlookFolder {
    param (
        [object]$Outlook,
        [string]$FolderType = "6" # Default to Inbox
    )
    $namespace = $Outlook.GetNamespace("MAPI")
    $folder = $namespace.GetDefaultFolder($FolderType)
    $explorer = $folder.GetExplorer()
    $explorer.Display()
}

function Open-EmailTemplate {
    param (
        [object]$Outlook,
        [string]$TemplatePath,
        [hashtable]$Replacements
    )
    $msg = $Outlook.CreateItemFromTemplate($TemplatePath)
    foreach ($key in $Replacements.Keys) {
        $msg.HTMLBody = $msg.HTMLBody.Replace($key, $Replacements[$key])
    }
    $inspector = $msg.GetInspector()
    $inspector.Display()
}

function Open-Emails {
    param (
        [datetime]$Today = (Get-Date)
    )
    $base = 'C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\'
    $emails = "$base\Email sablonok\"
    $outlook = New-Object -ComObject outlook.application

    # Open the Outlook folder
    Open-OutlookFolder -Outlook $outlook -FolderType "6"

    # Open BEVLÉT email
    $bevletTemplate = "$emails\BEVLÉT.oft"
    Open-EmailTemplate -Outlook $outlook -TemplatePath $bevletTemplate -Replacements @{
        "2023.??.??." = ($Today.AddDays(-1)).ToString("yyyy.MM.dd.")
    }

    # Open Bérleteken fennmaradt alkalmak email
    $berletTemplate = "$emails\Bérleteken fennmaradt alkalmak.oft"
    Open-EmailTemplate -Outlook $outlook -TemplatePath $berletTemplate -Replacements @{
        "2023.??.??." = ($Today.ToString("yyyy.MM.dd.") + " nyitás")
    }

    # Open DIÁKOK email
    $diakokTemplate = "$emails\DIÁKOK.oft"
    Open-EmailTemplate -Outlook $outlook -TemplatePath $diakokTemplate -Replacements @{}
}