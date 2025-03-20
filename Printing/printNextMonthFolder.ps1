Import-Module "$PSScriptRoot\MonthlyCommutingAllowancePrint.psm1"
Import-Module "$PSScriptRoot\monthlyPrintAttendceSheetCollegues.psm1"
Import-Module "$PSScriptRoot\monthlyPrintScheduleRequest.psm1"

$path = Get-SheetPath

Print-RequestFrontOffice
Print-AttandanceSheetFrontOffice -OpenFile $path

Print-RequestUszomester
Print-CommutingAllowance
Print-AttandanceSheetUszomester -OpenFile $path

Print-AttandanceSheetKarbantarto -OpenFile $path
Print-AttandanceSheetGepesz -OpenFile $path