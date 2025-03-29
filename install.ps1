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
function Create-ScheduledTask {
    param (
        [string]$TaskName,
        [string]$ScriptPath,
        [array]$Triggers,
        [string]$Description,
        [object]$Principal
    )
    $action = New-ScheduledTaskAction -Execute "C:\Program Files\PowerShell\7\pwsh.exe" -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$ScriptPath`""
    Register-ScheduledTask -TaskName $TaskName -Trigger $Triggers -Principal $Principal -Action $action -Description $Description
}

# Function to create registry keys
function Create-RegistryKeys {
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
$registryPath = "HKCU:\Software\Script1"
$registryProperties = @{
    EmailLastShown        = 0
    LastRun               = 0
    YesterdaysWorkingHours = 0
}

# Create registry keys
Create-RegistryKeys -RegistryPath $registryPath -Properties $registryProperties

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
        TaskName    = "ToggleTheme"
        ScriptPath  = "$PSScriptRoot\Theming\applyTheme.ps1"
        Triggers    = @()
        Description = "Runs the ToggleTheme script manually or when triggered by other means."
    },
    @{
        TaskName    = "Monthly print TIG"
        ScriptPath  = "$PSScriptRoot\Printing\MonthlyTIGprint.ps1"
        Triggers    = @(
            New-ScheduledTaskTrigger -Once -At (Get-Date "2025-03-31 15:30:00"),
            New-ScheduledTaskTrigger -Once -At (Get-Date "2025-04-30 15:30:00"),
            New-ScheduledTaskTrigger -Once -At (Get-Date "2025-05-31 15:30:00"),
            New-ScheduledTaskTrigger -Once -At (Get-Date "2025-06-30 15:30:00"),
            New-ScheduledTaskTrigger -Once -At (Get-Date "2025-08-31 15:30:00"),
            New-ScheduledTaskTrigger -Once -At (Get-Date "2025-09-30 15:30:00"),
            New-ScheduledTaskTrigger -Once -At (Get-Date "2025-10-31 15:30:00"),
            New-ScheduledTaskTrigger -Once -At (Get-Date "2025-11-30 15:30:00")
        )
        Description = "Runs the Monthly TIG print script."
    },
    @{
        TaskName    = "PrintWeeklyUNP"
        ScriptPath  = "$PSScriptRoot\Printing.Printing\weeklyUNPprint.ps1"
        Triggers    = @(New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At 3:00PM -EndBoundary "2025-06-13T23:59:59")
        Description = "Runs the printweeklyUNP script, which saves and prints sheets for UNP."
    },
    @{
        TaskName    = "Morning routine"
        ScriptPath  = "$PSScriptRoot\popUpShiftManager.ps1"
        Triggers    = @(New-ScheduledTaskTrigger -AtLogOn)
        Description = "Runs the Morning routine."
    },
    @{
        TaskName    = "PopUpMuszi"
        ScriptPath  = "$PSScriptRoot\popUpShiftManager.ps1"
        Triggers    = @(
            New-ScheduledTaskTrigger -Daily -At 12:58,
            New-ScheduledTaskTrigger -Daily -At 13:00,
            New-ScheduledTaskTrigger -AtLogOn
        )
        Description = "Runs the PopUpMuszi script manually or when triggered by other means."
    }
)

# Register all scheduled tasks
foreach ($task in $tasks) {
    Create-ScheduledTask -TaskName $task.TaskName -ScriptPath $task.ScriptPath -Triggers $task.Triggers -Description $task.Description -Principal $principal
}

# Create a script on the desktop
$modulePathTicket = Join-Path -Path $PSScriptRoot -ChildPath "maintenanceTicket.psm1"
$modulePathRoster = Join-Path -Path $PSScriptRoot -ChildPath "RosterInformation.psm1"
$newScriptContent = @"
Import-Module "$($modulePathTicket)"
Import-Module "$($modulePathRoster)"

Create-Ticket
"@
Create-ScriptOnDesktop -FileName "Hibajegy.ps1" -Content $newScriptContent