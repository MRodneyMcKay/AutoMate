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

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO",
        [switch]$ShowMessageBox,
        [string]$MsgBoxTitle = "HIBA",
        [string]$MsgBoxMessage
    )

    # Create timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"

    # Create ASCII Art for "AutoMate"
    $asciiArt = Get-AsciiArt

    # Build log path
    $date = Get-Date -Format "MM-dd"

    $logPath = Join-Path -Path $PSScriptRoot -ChildPath "..\..\..\logs\"

    if (-not (Test-Path $logPath)) {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
    }

    $resolvedLogPath = (Resolve-Path -Path $logPath).Path

    $logFile = Join-Path -Path $resolvedLogPath -ChildPath "$date.log"

    # Get script and function details from call stack
    $callStack = Get-PSCallStack

    $scriptName = $callStack[1].ScriptName
    $functionName = $callStack[1].FunctionName
    $lineNumber = $callStack[1].ScriptLineNumber

    # If running interactively, use defaults
    if (-not $scriptName) {
        $scriptName = "InteractiveShell"
    } else {
        $scriptName = [System.IO.Path]::GetFileName($scriptName)
    }

    if (-not $functionName) {
        $functionName = "GlobalScope"
    }

    if (-not $lineNumber) {
        $lineNumber = "Unknown"
    }

    # Thread / Job information
    $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId

    $jobName = "Main"

    try {
        if ($PSPrivateMetadata.JobName) {
            $jobName = $PSPrivateMetadata.JobName
        }
    } catch {
    }

    # Format log entry
    $logEntry = "[$timestamp] [$Level] [Job: $jobName] [Thread: $threadId] [$scriptName] [$functionName] [Line $lineNumber] $Message"

    # Create mutex for thread-safe logging
    $mutexName = "Global\AutoMateLogMutex"

    $mutex = [System.Threading.Mutex]::new($false, $mutexName)

    try {

        # Wait for exclusive access
        $null = $mutex.WaitOne()

        # Ensure log file exists
        if (-not (Test-Path $logFile)) {

            $asciiArt | Out-File -FilePath $logFile -Encoding UTF8
        }

        # Append log entry safely
        Add-Content -Path $logFile -Value $logEntry

    } finally {

        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }

    # Write to console with colors
    switch ($Level) {
        "INFO"    { Write-Host $logEntry -ForegroundColor Cyan }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "DEBUG"   { Write-Host $logEntry -ForegroundColor Gray }
    }

    # Map log level → WPF icon
    $icon = switch ($Level) {
        "INFO" { [System.Windows.MessageBoxImage]::Information }
        "WARNING" { [System.Windows.MessageBoxImage]::Warning }
        "ERROR" { [System.Windows.MessageBoxImage]::Error }
        "DEBUG" { [System.Windows.MessageBoxImage]::None }
    }

    # Show message box if requested
    if ($ShowMessageBox) {

        $finalTitle = $MsgBoxTitle

        if (-not $finalTitle) {
            $finalTitle = "$scriptName : $functionName"
        }

        $finalMessage = $MsgBoxMessage

        if (-not $finalMessage) {
            $finalMessage = $Message
        }

        Show-Message `
            -Message $finalMessage `
            -Title $finalTitle `
            -Icon $icon
    }
}