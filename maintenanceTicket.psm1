Import-Module $PSScriptRoot\RosterInformation.psm1

function Create-Ticket {
    
$shiftManager = (Get-ShiftManager)[(Get-Date).TimeOfDay -lt (New-TimeSpan -Hours 12 -Minutes 55) ? 0 : 1].Name


$today = Get-Date
$Excel = New-Object -ComObject Excel.Application
$Excel.visible=$false

$outlook = New-Object -comObject Outlook.Application
$mail = $outlook.createitemfromtemplate("C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Email sablonok\HIBA.oft")
$inspector = $mail.GetInspector
$inspector.Display()

$Workbook = $Excel.Workbooks.Open("C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\HIBA\-=SABLON=-\HIBA jelentés.xlsx")
$workSheet = $Workbook.Sheets.Item("Munka1")
$workSheet.Range("A6").value()= Get-Date $today -Format "yyyy.MM.dd. HH:mm"
$workSheet.Range("A23").value()= $shiftManager
$Excel.visible=$true
}