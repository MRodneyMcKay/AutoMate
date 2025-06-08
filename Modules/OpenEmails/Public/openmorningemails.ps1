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

function Open-Emails {
    param (
        [datetime]$Today = (Get-Date)
    )
    $base = 'C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\'
    $emails = "$base\Email sablonok\"
    Start-Process "OUTLOOK"
    Write-Log -Message "Starting Outlook..." -Level "INFO"
    # Wait for Outlook to initialize by checking every second
    $maxWaitTime = 30 # Maximum wait time in seconds (5 minutes)
    $waitedTime = 0
    $outlook = $null

    while ($waitedTime -lt $maxWaitTime) {
        try {
            $outlook = [InteropCom]::GetActiveInstance("Outlook.Application", $true)
            if ($outlook -and $outlook.Session) {
                Write-Log -Message "Outlook is ready." -Level "INFO"
                break
            }
        } catch {
            # Ignore errors during the wait period
            Write-Log -Message "Waiting for Outlook to be ready..." -Level "INFO"
        }
        Start-Sleep -Seconds 1
        $waitedTime++
    }

    # If Outlook is still not ready, explicitly create a COM object and wait 10 seconds
    if (-not $outlook -or -not $outlook.Session) {
        Write-Output "Outlook not ready after $maxWaitTime seconds. Creating COM object explicitly."
        try {
            $outlook = New-Object -ComObject outlook.application
            Start-Sleep -Seconds 10 # Wait for initialization
        } catch {
            Write-Log -Message "Failed to create Outlook application. Error: $_" -Level "ERROR"
            return # Exit the function if Outlook cannot be created
        }
    }

    # Validate Outlook object
    if (-not $outlook -or -not $outlook.Session) {
        Write-Log -Message "Outlook COM object is not functional after creation." -Level "ERROR"
        return
    }

    try {
        # Open BEVLÉT email
        $bevletTemplate = "$emails\BEVLÉT.oft"
        Open-EmailTemplate -Outlook $outlook -TemplatePath $bevletTemplate -Replacements @{
            "2023.??.??." = ($Today.AddDays(-1)).ToString("yyyy.MM.dd.")
        } -subject "BEVLÉT"

        # Open Bérleteken fennmaradt alkalmak email
        $berletTemplate = "$emails\Bérleteken fennmaradt alkalmak.oft"
        Open-EmailTemplate -Outlook $outlook -TemplatePath $berletTemplate -Replacements @{
            "2023.??.??." = ($Today.ToString("yyyy.MM.dd.") + " nyitás")
        } -subject "Bérletes"

        # Open DIÁKOK email
        $diakokTemplate = "$emails\DIÁKOK.oft"
        Open-EmailTemplate -Outlook $outlook -TemplatePath $diakokTemplate -Replacements @{} -subject "Diákok"
        # Open karórák email only on Mondays
        if ($Today.DayOfWeek -eq 'Monday') {
            $karorakTemplate = "$emails\Órák.oft"
            Open-EmailTemplate -Outlook $outlook -TemplatePath $karorakTemplate -Replacements @{} -subject "Karórák"
        }
    } catch {
        Write-Log -Message "Error opening email templates: $_" -Level "ERROR"
    } finally {
        if ($outlook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)
            $outlook = $null
        }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}