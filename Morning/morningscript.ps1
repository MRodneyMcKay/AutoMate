$registryPath = "HKCU:\Software\Script"
if ($(Get-ItemPropertyValue -Path $registryPath -Name LastRun) -eq (Get-Date).Day) {
    exit
}

Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\RosterInformation.psm1')
Import-Module "$PSScriptRoot\createDir.psm1"
Import-Module "$PSScriptRoot\createWorkingHoursExcel.psm1"
Import-Module "$PSScriptRoot\openEmails.psm1"

create-Directories

if ($(Get-ItemPropertyValue -Path $registryPath -Name EmailLastShown) -ne (Get-Date).Day) {
    Open-Emails
    Set-ItemProperty -Path $registryPath -Name EmailLastShown -value (Get-Date).Day
}
if ($(Get-ItemPropertyValue -Path $registryPath -Name YesterdaysWorkingHours) -ne (Get-Date).Day) {
    $roster = Get-Receptionists
    $approved = @(1, 112, "1/9")
    open-StudentWorkReportFurdo -fillCompletely ((($roster | Where-Object { $_.Name -eq 'Raduska Zsolt' }).Shift -in $approved -or ($roster | Where-Object { $_.Name -eq 'Konfár Nikolett' }).Shift -in $approved))
    Set-ItemProperty -Path $registryPath -Name YesterdaysWorkingHours -value (Get-Date).Day
}

Set-ItemProperty -Path $registryPath -Name LastRun -value (Get-Date).Day