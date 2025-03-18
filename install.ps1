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
$scriptPath = "$PSScriptRoot\Printing.Printing\weeklyUNPprint.ps1"

# Define the action to run the PowerShell script
$action = New-ScheduledTaskAction -Execute "C:\Program Files\PowerShell\7\pwsh.exe" -Argument "-WindowStyle hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

# Create a trigger to execute the task every Friday at 15:00 until 06/13
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At 3:00PM -EndBoundary "2025-06-13T23:59:59"

# Register the scheduled task without a trigger
Register-ScheduledTask -TaskName "PrintWeeklyUNP"  -Trigger $trigger -Principal $principal -Action $action -Description "Runs the printweeklyUNP script, which saves and prints shhets for UNP."