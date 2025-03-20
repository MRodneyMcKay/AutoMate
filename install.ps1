# Define the path to your PowerShell script
$scriptPath = "$PSScriptRoot\Theming\themeChangeManager.ps1"

# Create a trigger to execute the task at login
$trigger = New-ScheduledTaskTrigger -AtLogOn

# Define the action to run the PowerShell script
$action = New-ScheduledTaskAction -Execute "C:\Program Files\PowerShell\7\pwsh.exe" -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
# Define a principal that runs the task with highest privileges
$principal = New-ScheduledTaskPrincipal -UserId "HIROSSPORT" -LogonType ServiceAccount -RunLevel Highest

# Register the scheduled task
Register-ScheduledTask -TaskName "Theme Change Manager" -Trigger $trigger -Principal $principal -Action $action -Description "Runs the Theme Change Manager script at user login."

# Define the path to your PowerShell script
$scriptPath = "$PSScriptRoot\Theming\applyTheme.ps1"

# Define the action to run the PowerShell script
$action = New-ScheduledTaskAction -Execute "C:\Program Files\PowerShell\7\pwsh.exe" -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

# Register the scheduled task without a trigger
Register-ScheduledTask -TaskName "ToggleTheme" -Principal $principal -Action $action -Description "Runs the ToggleTheme script manually or when triggered by other means."

# Define the path to your PowerShell script
$scriptPath = "$PSScriptRoot\Printing\MonthlyTIGprint.ps1"

# Create a trigger to execute the task at login

$trigger1 = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-03-31 15:30:00")
$trigger2 = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-04-30 15:30:00")
$trigger3 = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-05-31 15:30:00")
$trigger4 = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-06-30 15:30:00")
$trigger5 = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-08-31 15:30:00")
$trigger6 = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-09-30 15:30:00")
$trigger7 = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-10-31 15:30:00")
$trigger8 = New-ScheduledTaskTrigger -Once -At (Get-Date "2025-11-30 15:30:00")

# Combine all the triggers
$triggers = @($trigger1, $trigger2, $trigger3, $trigger4, $trigger5, $trigger6, $trigger7, $trigger8)

# Define the action to run the PowerShell script
$action = New-ScheduledTaskAction -Execute "C:\Program Files\PowerShell\7\pwsh.exe" -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
# Define a principal that runs the task with highest privileges
$principal = New-ScheduledTaskPrincipal -UserId "HIROSSPORT" -LogonType ServiceAccount -RunLevel Highest

# Register the scheduled task
Register-ScheduledTask -TaskName "Monthly print TIG" -Trigger $triggers -Principal $principal -Action $action -Description "Runs the Monthly TIG print script"

$scriptPath = "$PSScriptRoot\Printing.Printing\weeklyUNPprint.ps1"

# Define the action to run the PowerShell script
$action = New-ScheduledTaskAction -Execute "C:\Program Files\PowerShell\7\pwsh.exe" -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

# Create a trigger to execute the task every Friday at 15:00 until 06/13
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At 3:00PM -EndBoundary "2025-06-13T23:59:59"

# Register the scheduled task without a trigger
Register-ScheduledTask -TaskName "PrintWeeklyUNP"  -Trigger $trigger -Principal $principal -Action $action -Description "Runs the printweeklyUNP script, which saves and prints shhets for UNP."

# Define the content of the new script
$modulePathTicket = Join-Path -Path $PSScriptRoot -ChildPath "maintenanceTicket.psm1"
$modulePathRoster = Join-Path -Path $PSScriptRoot -ChildPath "RosterInformation.psm1"
$newScriptContent = @"
Import-Module "$($modulePathTicket)"
Import-Module "$($modulePathRoster)"

Create-Ticket
"@

# Get the Desktop path
$desktopPath = [Environment]::GetFolderPath("Desktop")

# Define the file name and path for the new script
$newScriptPath = Join-Path -Path $desktopPath -ChildPath "Hibajegy.ps1"

# Create the script file and write the content to it
Set-Content -Path $newScriptPath -Value $newScriptContent

# Define the path to your PowerShell script
$scriptPath = "$PSScriptRoot\popUpShiftManager.ps1"

# Define the action to run the PowerShell script
$action = New-ScheduledTaskAction -Execute "C:\Program Files\PowerShell\7\pwsh.exe" -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

$trigger1 = New-ScheduledTaskTrigger -AtLogOn
$trigger2 = New-ScheduledTaskTrigger -Daily -At 12:58
$trigger3 = New-ScheduledTaskTrigger -Daily -At 13:00
# Combine all the triggers
$triggers = @($trigger1, $trigger2, $trigger3)


# Register the scheduled task without a trigger
Register-ScheduledTask -TaskName "PopUpMuszi" -Trigger $triggers -Principal $principal -Action $action -Description "Runs the ToggleTheme script manually or when triggered by other means."
