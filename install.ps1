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

# Define reusable functions for common tasks

# Function to create a scheduled task
function New-ScheduledTask {
    param (
        [string]$TaskName,
        [string]$ScriptPath,
        [array]$Triggers,
        [string]$Description,
        [object]$Principal
    )
    $pwshPath = (Get-Command pwsh).Source
    if (-not (Test-Path $pwshPath)) {
        throw "PowerShell executable not found."
    }
    $action = New-ScheduledTaskAction -Execute $pwshPath -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$ScriptPath`""
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }
    Register-ScheduledTask -TaskName $TaskName -Principal $Principal -Action $action -Description $Description
    Set-ScheduledTask -TaskName $TaskName -Trigger $triggers
}

# Function to create registry keys
function New-RegistryKeys {
    param (
        [string]$RegistryPath,
        [hashtable]$Properties
    )
    New-Item -Path $RegistryPath -Force
    foreach ($key in $Properties.Keys) {
        Set-ItemProperty -Path $RegistryPath -Name $key -Value $Properties[$key]
    }
}

# Function to create a PowerShell script on the desktop
function Create-ScriptOnDesktop {
    param (
        [string]$FileName,
        [string]$Content
    )
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $scriptPath = Join-Path -Path $desktopPath -ChildPath $FileName
    Set-Content -Path $scriptPath -Value $Content
}

# Define the principal for scheduled tasks
$principal = New-ScheduledTaskPrincipal -UserId "HIROSSPORT" -LogonType ServiceAccount -RunLevel Highest

# Define registry path and properties
$registryPath = "HKCU:\Software\Script"
$registryProperties = @{
    EmailLastShown        = 0
    LastRun               = 0
    YesterdaysWorkingHours = 0
}

# Create registry keys
New-RegistryKeys -RegistryPath $registryPath -Properties $registryProperties

# Define scheduled tasks
$tasks = @(
    @{
        TaskName    = "Open UNP sheet"
        ScriptPath  = "$PSScriptRoot\Morning\openUNPsheet.ps1"
        Triggers    = @(New-ScheduledTaskTrigger -Daily -At 7:30)
        Description = "Runs the Theme Change Manager script at user login."
    },
    @{
        TaskName    = "Theme Change Manager"
        ScriptPath  = "$PSScriptRoot\Theming\themeChangeManager.ps1"
        Triggers    = @(New-ScheduledTaskTrigger -AtLogOn)
        Description = "Runs the Theme Change Manager script at user login."
    },    
    @{
        TaskName    = "Monthly print TIG"
        ScriptPath  = "$PSScriptRoot\Printing\MonthlyTIGprint.ps1"
        Triggers    = @(
            New-ScheduledTaskTrigger -Once -At ([datetime] "2025-03-31 15:30:00")
            New-ScheduledTaskTrigger -Once -At ([datetime] "2025-04-30 15:30:00")
            New-ScheduledTaskTrigger -Once -At ([datetime] "2025-05-31 15:30:00")
            New-ScheduledTaskTrigger -Once -At ([datetime] "2025-06-30 15:30:00")
            New-ScheduledTaskTrigger -Once -At ([datetime] "2025-08-31 15:30:00")
            New-ScheduledTaskTrigger -Once -At ([datetime] "2025-09-30 15:30:00")
            New-ScheduledTaskTrigger -Once -At ([datetime] "2025-10-31 15:30:00")
            New-ScheduledTaskTrigger -Once -At ([datetime] "2025-11-30 15:30:00")
        )
        Description = "Runs the Monthly TIG print script."
    },
    @{
        TaskName    = "PrintWeeklyUNP"
        ScriptPath  = "$PSScriptRoot\Printing\weeklyUNPprint.ps1"
        Triggers    = @(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At (Get-Date -Hour 15 -Minute 0))
        Description = "Runs the printweeklyUNP script, which saves and prints sheets for UNP."
    },
    @{
        TaskName    = "Morning routine"
        ScriptPath  = "$PSScriptRoot\Morning\morningscript.ps1"
        Triggers    = @(New-ScheduledTaskTrigger -AtLogOn)
        Description = "Runs the Morning routine."
    },
    @{
        TaskName    = "PopUpMuszi"
        ScriptPath  = "$PSScriptRoot\popUpShiftManager.ps1"
        Triggers    = @(
            New-ScheduledTaskTrigger -Daily -At (Get-Date -Hour 12 -Minute 58)
            New-ScheduledTaskTrigger -Daily -At (Get-Date -Hour 13 -Minute 0)
            New-ScheduledTaskTrigger -AtLogOn
        )
        Description = "Runs the PopUpMuszi script manually or when triggered by other means."
    },
    @{
        TaskName    = "ToggleTheme"
        ScriptPath  = "$PSScriptRoot\Theming\applyTheme.ps1"
        Triggers    = @()
        Description = "Runs the ToggleTheme script manually or when triggered by other means."
    }
)

# Register all scheduled tasks
foreach ($task in $tasks) {
    if ($task.TaskName -eq "PrintWeeklyUNP"){
        $task.Triggers[0].EndBoundary = ([datetime] "2025-06-30 15:00:00")
    }    
    New-ScheduledTask -TaskName $task.TaskName -ScriptPath $task.ScriptPath -Triggers $task.Triggers -Description $task.Description -Principal $principal
}

# Create a script on the desktop
$modulePathTicket = Join-Path -Path $PSScriptRoot -ChildPath "maintenanceTicket.psm1"
$newScriptContent = @"
Import-Module "$($modulePathTicket)"

New-Ticket
"@
Create-ScriptOnDesktop -FileName "Hibajegy.ps1" -Content $newScriptContent

#saving office interop dll assembly paths to environment variables
# Define Office Interop assemblies
$assemblies = @("Excel", "Outlook", "Word")

# Iterate over each assembly, load it, get the location, and store it in environment variables
foreach ($assembly in $assemblies) {
    $dll_path = powershell.exe -Command "[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.$assembly') | Select-Object -ExpandProperty Location | Out-String"
    [System.Environment]::SetEnvironmentVariable("OfficeAssemblies_$assembly", $dll_path.Trim(), [System.EnvironmentVariableTarget]::User)
}