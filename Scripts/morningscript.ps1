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

$registryPath = "HKCU:\Software\Script"

$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$resolvedModulegPath = (Resolve-Path -Path $ModulePath).Path

# Load only what main thread needs
Import-Module (Join-Path $resolvedModulegPath 'LoggingSystem\LoggingSystem.psd1')
Import-Module (Join-Path $resolvedModulegPath 'CreateMaintanenceDir\CreateMaintanenceDir.psd1')

Import-Module ThreadJob

# THREAD 01 - EMAILS
$job1 = Start-ThreadJob -Name "Emails" -ScriptBlock {
    param($registryPath, $resolvedModulegPath)

    Import-Module (Join-Path $resolvedModulegPath 'LoggingSystem\LoggingSystem.psd1')
    Import-Module (Join-Path $resolvedModulegPath 'OpenEmails\OpenEmails.psd1')

    if ((Get-ItemPropertyValue -Path $registryPath -Name EmailLastShown) -ne (Get-Date).Day) {
        Write-Log -Message "Emailek megnyitása"
        Open-Emails
        Set-ItemProperty -Path $registryPath -Name EmailLastShown -value (Get-Date).Day
    } else {
        Write-Log -Message "Emailek már megnyitva voltak"
    }
} -ArgumentList $registryPath, $resolvedModulegPath


# THREAD 02 - STUDENT REPORT
$job2 = Start-ThreadJob -Name "StudentReport" -ScriptBlock {
    param($registryPath, $resolvedModulegPath)

    Import-Module (Join-Path $resolvedModulegPath 'LoggingSystem\LoggingSystem.psd1')
    Import-Module (Join-Path $resolvedModulegPath 'RosterInformation\RosterInformation.psd1')
    Import-Module (Join-Path $resolvedModulegPath 'StudentWorkReport\StudentWorkReport.psd1')

    if ((Get-ItemPropertyValue -Path $registryPath -Name YesterdaysWorkingHours) -ne (Get-Date).Day) {
        Write-Log -Message "Diákelszámolás elkészítése"

        $roster = Get-Receptionists
        $approved = @(1, 112, "1/9")

        $fill = (
            ($roster | Where-Object Name -eq 'Raduska Zsolt').Shift -in $approved -or
            ($roster | Where-Object Name -eq 'Konfár Nikolett').Shift -in $approved -or
            ($roster | Where-Object Name -eq 'Antal Natália').Shift -in $approved
        )

        open-StudentWorkReportFurdo -fillCompletely $fill

        if ((Get-Date).Month -in 6,7,8,9) {
            open-StudentWorkReportStrand -fillCompletely $fill
        }

        Set-ItemProperty -Path $registryPath -Name YesterdaysWorkingHours -value (Get-Date).Day
    } else {
        Write-Log -Message "Diákelszámolás már elkészült"
    }
} -ArgumentList $registryPath, $resolvedModulegPath


# THREAD 03 - LANE PRINTING
$job3 = Start-ThreadJob -Name "LanePrint" -ScriptBlock {
    param($registryPath, $resolvedModulegPath)

    Import-Module (Join-Path $resolvedModulegPath 'LoggingSystem\LoggingSystem.psd1')
    Import-Module (Join-Path $resolvedModulegPath 'LaneOccupancy\LaneOccupancy.psd1')

    if ((Get-ItemPropertyValue -Path $registryPath -Name LaneSchedulePrinted) -ne (Get-Date).Day) {
        Write-Log -Message "Pályabeosztás nyomtatása"
        Print-Today
        Set-ItemProperty -Path $registryPath -Name LaneSchedulePrinted -value (Get-Date).Day
    } else {
        Write-Log -Message "Pályabeosztás már nyomtatva volt"
    }
} -ArgumentList $registryPath, $resolvedModulegPath


# WAIT FOR ALL THREADS
$jobs = @($job1, $job2, $job3)

$jobs | Wait-Job

# Collect output + errors
foreach ($job in $jobs) {
    Receive-Job $job -ErrorAction SilentlyContinue
}

# Cleanup
$jobs | Remove-Job

# FINAL STEP (AFTER THREADS)
create-Directories
Start-Process msedge
Set-ItemProperty -Path $registryPath -Name LastRun -value (Get-Date).Day