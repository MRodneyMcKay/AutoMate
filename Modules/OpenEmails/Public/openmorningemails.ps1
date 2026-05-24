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

    # Wait for Outlook COM initialization
    $maxWaitTime = 30
    $waitedTime = 0
    $outlook = $null

    while ($waitedTime -lt $maxWaitTime) {
        try {
            $outlook = [InteropCom]::GetActiveInstance("Outlook.Application", $true)

            # Session existence alone is not enough,
            # but it is enough for initial startup validation
            if ($outlook -and $outlook.Session) {
                Write-Log -Message "Outlook COM initialized." -Level "INFO"
                break
            }
        } catch {
            Write-Log -Message "Waiting for Outlook COM..." -Level "INFO"
        }

        Start-Sleep -Seconds 1
        $waitedTime++
    }

    # Fallback COM creation
    if (-not $outlook -or -not $outlook.Session) {
        Write-Log -Message "Outlook not ready after $maxWaitTime seconds. Creating COM object explicitly." -Level "WARNING"

        try {
            $outlook = New-Object -ComObject outlook.application
            Start-Sleep -Seconds 10
        } catch {
            Write-Log -Message "Failed to create Outlook application. Error: $_" -Level "ERROR"
            return
        }
    }

    # Final validation
    if (-not $outlook -or -not $outlook.Session) {
        Write-Log -Message "Outlook COM object is not functional after creation." -Level "ERROR"
        return
    }

    # Build email queue
    $emailQueue = @()

    $emailQueue += @{
        TemplatePath = "$emails\BEVLÉT.oft"
        Replacements = @{
            "2023.??.??." = ($Today.AddDays(-1)).ToString("yyyy.MM.dd.")
        }
        Subject = "BEVLÉT"
    }

    $emailQueue += @{
        TemplatePath = "$emails\Bérleteken fennmaradt alkalmak.oft"
        Replacements = @{
            "2023.??.??." = ($Today.ToString("yyyy.MM.dd.") + " nyitás")
        }
        Subject = "Bérletes"
    }

    $emailQueue += @{
        TemplatePath = "$emails\DIÁKOK.oft"
        Replacements = @{}
        Subject = "Diákok"
    }

    if ($Today.DayOfWeek -eq 'Tuesday') {
        $emailQueue += @{
            TemplatePath = "$emails\Órák.oft"
            Replacements = @{}
            Subject = "Karórák"
        }
    }

    # Process email queue
    try {

        foreach ($email in $emailQueue) {

            while ($true) {

                try {

                    Write-Log -Message "Opening email: $($email.Subject)" -Level "INFO"

                    Open-EmailTemplate `
                        -Outlook $outlook `
                        -TemplatePath $email.TemplatePath `
                        -Replacements $email.Replacements `
                        -subject $email.Subject

                    Write-Log -Message "Successfully opened: $($email.Subject)" -Level "INFO"

                    break

                } catch {

                    $message = $_.Exception.Message

                    # Outlook modal dialog / busy state
                    if (
                        $message -like "*párbeszédpanel*" -or
                        $message -like "*dialog box*" -or
                        $message -like "*cannot perform this action*" -or
                        $message -like "*another program is using Outlook*"
                    ) {

                        Write-Log -Message "Outlook is blocked by a dialog. Waiting before retry..." -Level "WARNING"

                        Start-Sleep -Seconds 2

                        continue
                    }

                    # Unknown error
                    throw
                }
            }
        }

    } catch {
        Write-Log -Message "Error opening email templates: $_" -Level "ERROR" -ShowMessageBox
    } finally {

        if ($outlook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
            $outlook = $null
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}