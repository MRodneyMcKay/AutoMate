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
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'RosterInformation\RosterInformation.psd1')
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'Themes\themes.psd1')

try {
$roster = Get-Receptionists
} catch {
    Set-Default
    exit
}
$zsoltiToday = ($roster | Where-Object { $_.Name -eq 'Raduska Zsolt' }).Shift
$nikiToday = ($roster | Where-Object { $_.Name -eq 'Konfár Nikolett' }).Shift
# Get date and working states


# Define common function for logic
function ProcessShift {
    param (
        [string]$Shift,
        [datetime]$TriggerTime,
        [scriptblock]$Action,
        [string]$applyThemeToggle
    )
    $DateTimeNow = Get-Date
    if ($DateTimeNow.TimeOfDay -lt $TriggerTime.TimeOfDay) {
        & $Action
    }
    [System.Environment]::SetEnvironmentVariable("ApplyTheme", $applyThemeToggle, [System.EnvironmentVariableTarget]::User)
    $triggers = New-ScheduledTaskTrigger -At $TriggerTime -Once
    try {
        Set-ScheduledTask -TaskName "ToggleTheme" -Trigger $triggers -ErrorAction Stop
        Write-Log -Message "Scheduled task created at $TriggerTime"
    } catch {
        Write-Log -Message "Failed to set scheduled task: $_" -Level WARNING
    }
    
}

# Define task and action mappings for shifts
$shiftConfigsZsolti = @(
    @{ Shift = 1; TriggerTime = ($nikiToday -eq 2 ? [datetime]"13:00" : [datetime]"13:30"); Action = { Set-Zsolti}; applyThemeToggle = ($nikiToday -eq 2 ? "Niki" : "Default")},
    @{ Shift = 2; TriggerTime = [datetime]"13:00"; Action = ($nikiToday -eq 1 ? {Set-Niki} : {Set-Default}); applyThemeToggle = "Zsolti"},
    @{ Shift = 112; TriggerTime = [datetime]"17:30"; Action = { Set-Zsolti }; applyThemeToggle = "Default" },
    @{ Shift = 212; TriggerTime = [datetime]"08:30"; Action = { Set-Default }; applyThemeToggle = "Zsolti" },
    @{ Shift = "1/9"; TriggerTime = [datetime]"14:30"; Action = { Set-Zsolti }; applyThemeToggle = "Default" },
    @{ Shift = "2/9"; TriggerTime = [datetime]"12:00"; Action = { Set-Default }; applyThemeToggle = "Zsolti" }
)

$shiftConfigsNiki = @(
    @{ Shift = 1; TriggerTime = ($zsoltiToday -eq 2 ? [datetime]"13:00" : [datetime]"13:30"); Action = { Set-Niki}; applyThemeToggle = ($zsoltiToday -eq 2 ? "Zsolti" : "Default")},
    @{ Shift = 2; TriggerTime = [datetime]"13:00"; Action = ($zsoltiToday -eq 1 ? {Set-Zsolti} : {Set-Default}); applyThemeToggle = "Niki"},
    @{ Shift = 112; TriggerTime = [datetime]"17:30"; Action = { set-Niki }; applyThemeToggle = "Default" },
    @{ Shift = 212; TriggerTime = [datetime]"08:30"; Action = { Set-Default }; applyThemeToggle = "Niki" },
    @{ Shift = "1/9"; TriggerTime = [datetime]"14:30"; Action = { Set-Niki }; applyThemeToggle = "Default" },
    @{ Shift = "2/9"; TriggerTime = [datetime]"12:00"; Action = { Set-Default }; applyThemeToggle = "Niki" }
)

# Function to configure shifts
function ConfigureShifts {
    param (
        [array]$ShiftConfigs,
        [string]$todayShift
    )

    foreach ($config in $ShiftConfigs) {
        if ($todayShift -eq $config.Shift) {
            ProcessShift $config.Shift $config.TriggerTime $config.Action $config.applyThemeToggle
            return $true
        }
    }
    return $false
}

# Process shifts
$managedZsolti = ConfigureShifts $shiftConfigsZsolti $zsoltiToday
$managedNiki = ConfigureShifts $shiftConfigsNiki $nikiToday

# Final fallback action
if (-not ($managedZsolti -or $managedNiki)) {
    Set-Default
}