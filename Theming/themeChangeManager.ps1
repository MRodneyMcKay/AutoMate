Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\RosterInformation.psm1')
Import-Module $PSScriptRoot\themes.psm1

$roster = Get-Receptionists
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
    Set-ScheduledTask -TaskName "ToggleTheme" -Trigger $triggers -Verbose
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
    @{ Shift = 1; TriggerTime = ($nikiToday -eq 2 ? [datetime]"13:00" : [datetime]"13:30"); Action = { Set-Niki}; applyThemeToggle = ($zsoltiToday -eq 2 ? "Zsolti" : "Default")},
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