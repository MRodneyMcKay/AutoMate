Import-Module "$PSScriptRoot\themes.psm1"

# Get the value of the environment variable
$theme = [System.Environment]::GetEnvironmentVariable('ApplyTheme', [System.EnvironmentVariableTarget]::User)

# Map the user choice to the function name and invoke the function using the call operator (&)
$functionName = "Set-$theme"
if (Get-Command $functionName -ErrorAction SilentlyContinue) {
    & $functionName
} else {
    Write-Error "Unknown environmental variable value"
}
