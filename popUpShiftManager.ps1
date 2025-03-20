Import-Module $PSScriptRoot\RosterInformation.psm1

[System.Reflection.Assembly]::LoadFrom("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll")
Add-Type -AssemblyName PresentationFramework

$shiftManager =  (Get-ShiftManager)[(Get-Date).TimeOfDay -lt (New-TimeSpan -Hours 12 -Minutes 55) ? 0 : 1]

if ($null -eq $shiftManager.Name) {
    [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "Nem tudom ki a $(($shiftManager.Shift -eq "1~*") ? "délelöttös" : "délutános") műszakvezető", "Vajon ki a müszi?", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Error)
}
else {
    [System.Windows.Forms.MessageBox]::Show($(New-Object -TypeName System.Windows.Forms.Form -Property @{TopMost=$true}), "A $(($shiftManager.Shift -eq "1*") ? "délelöttös" : "délutános") műszakvezető: $($shiftManager.Name)", "Vajon ki a müszi?", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.MessageBoxImage]::Information)
}