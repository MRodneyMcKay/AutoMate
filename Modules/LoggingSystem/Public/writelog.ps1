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

    # Create a timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Create ASCII Art for "AutoMate"
    $asciiArt = Get-AsciiArt
    $date =Get-Date -Format "MM-dd"
    $logPath = Join-Path -Path $PSScriptRoot -ChildPath "..\..\..\logs\"
    $resolvedLogPath = (Resolve-Path -Path $logPath).Path
    $logFile = Join-Path -Path $resolvedLogPath -ChildPath "$date.log"
    if (-Not (Test-Path $resolvedLogPath)) {
        New-Item -Path $resolvedLogPath -ItemType Directory -Force
    }    
    # Get script and function details from the call stack
    $callStack = Get-PSCallStack
    $scriptName = $callStack[1].ScriptName
    $functionName = $callStack[1].FunctionName
    $lineNumber = $callStack[1].ScriptLineNumber

    # If running interactively, use defaults
    if (-not $scriptName) { $scriptName = "InteractiveShell" } else {  $scriptName = [System.IO.Path]::GetFileName($scriptName)}
    if (-not $functionName) { $functionName = "GlobalScope" }
    if (-not $lineNumber) { $lineNumber = "Unknown" }

    # Format log entry
    $logEntry = "[$timestamp] [$Level] [$scriptName] [$functionName] [Line $lineNumber] $Message"

    # Ensure the log file exists and append the ASCII art and message
    if (-not (Test-Path $logFile)) {
        # Create log file and add ASCII Art header
        $asciiArt | Out-File -FilePath $logFile -Encoding UTF8
    }

     # Append to log file
     Add-Content -Path $LogFile -Value $logEntry

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
        $finalTitle = $MsgBoxTitle ? $MsgBoxTitle : "$scriptName : $functionName" 
        $finalMessage = $MsgBoxMessage ? $MsgBoxMessage : $Message 
        Show-Message -Message $finalMessage -Title $finalTitle -Icon $icon 
    }
}