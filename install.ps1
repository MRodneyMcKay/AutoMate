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

Import-Module $PSScriptRoot\log.psm1

# Function to create a scheduled task
function New-ScheduledTask {
    param (
        [string]$TaskName,
        [string]$ScriptPath,
        [array] $Triggers,
        [string]$Description,
        [object]$Principal
    )
    $pwshPath = (Get-Command pwsh).Source
    if (-not (Test-Path $pwshPath)) {
        Write-Log -Message "PowerShell executable not found." -Level "ERROR"
    }
    $action = New-ScheduledTaskAction -Execute $pwshPath -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$ScriptPath`""  -WorkingDirectory $PSScriptRoot
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }
    Register-ScheduledTask -TaskName $TaskName -Principal $Principal -Action $action -Description $Description
    Write-Log -Message "Scheduled task '$TaskName' created."
    Set-ScheduledTask -TaskName $TaskName -Trigger $triggers
    Write-Log -Message "Scheduled task '$TaskName' triggers set."

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
    Write-Log -Message "Registry keys created at '$RegistryPath'."
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
    Write-Log -Message "Script '$FileName' created on desktop."
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
$tasks = Get-Content .\scheduledTasks.json | ConvertFrom-Json

# Register all scheduled tasks
foreach ($task in $tasks) {
    # Prepare an array for scheduled task triggers
    $taskTriggers = @()

    # Process each trigger for the task
    foreach ($trigger in $task.Triggers) {
        if ($trigger.TriggerType -eq "Daily") {
            $taskTriggers += New-ScheduledTaskTrigger -Daily -At $trigger.Time
        }
        elseif ($trigger.TriggerType -eq "AtLogOn") {
            $taskTriggers += New-ScheduledTaskTrigger -AtLogOn
        }
        elseif ($trigger.TriggerType -eq "Once") {
            $taskTriggers += New-ScheduledTaskTrigger -Once -At ([datetime]$trigger.DateTime)
        }
        elseif ($trigger.TriggerType -eq "Weekly") {
            $taskTriggers += New-ScheduledTaskTrigger -Weekly -DaysOfWeek $trigger.DaysOfWeek -At $trigger.Time
        }
    }

    # Special condition for "PrintWeeklyUNP" task to set an end boundary for the trigger
    if ($task.TaskName -eq "PrintWeeklyUNP") {
        $taskTriggers[0].EndBoundary = ([datetime] "2025-06-30 15:00:00")
    }

    # Register the scheduled task
    New-ScheduledTask -TaskName $task.TaskName -ScriptPath $task.ScriptPath -Triggers $taskTriggers -Description $task.Description -Principal $principal
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
    Write-Log -Message "Office Interop assembly '$assembly' path saved to environment variable."
}