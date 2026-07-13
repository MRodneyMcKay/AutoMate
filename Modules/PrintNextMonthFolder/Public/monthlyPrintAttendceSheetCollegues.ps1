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

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Xml.Linq

# Path to the attendance XML file
$Script:AttendanceXmlPath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\nevek.xml"

# Function to load XML and get all names from a department, excluding a specific position
function Get-DepartmentNames {
    param (
        [string]$DepartmentName,
        [string]$ExcludePosition = "Önkormányzat"
    )
    
    try {
        $XDoc = [System.Xml.Linq.XDocument]::Load($Script:AttendanceXmlPath)
        $Root = $XDoc.Root
        
        $Dept = $Root.Elements("Department") | Where-Object { $_.Attribute("Name").Value -eq $DepartmentName } | Select-Object -First 1
        if (-not $Dept) {
            Write-Log -Message "Részleg nem találva: $DepartmentName" -Level "ERROR"
            return @()
        }
        
        $Names = @()
        foreach ($Position in $Dept.Elements("Position")) {
            $PosName = $Position.Attribute("Name").Value
            if ($PosName -ne $ExcludePosition) {
                foreach ($NameElement in $Position.Elements("Nev")) {
                    $Names += $NameElement.Value
                }
            }
        }
        
        return ($Names | Sort-Object) -as [string[]]
    }
    catch {
        Write-Log -Message "Hiba az XML olvasásakor: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

# Function to get only names from a specific position across all departments
function Get-PositionNames {
    param (
        [string]$PositionName = "Önkormányzat"
    )
    
    try {
        $XDoc = [System.Xml.Linq.XDocument]::Load($Script:AttendanceXmlPath)
        $Root = $XDoc.Root
        
        $Names = @()
        foreach ($Dept in $Root.Elements("Department")) {
            $Position = $Dept.Elements("Position") | Where-Object { $_.Attribute("Name").Value -eq $PositionName } | Select-Object -First 1
            if ($Position) {
                foreach ($NameElement in $Position.Elements("Nev")) {
                    $Names += $NameElement.Value
                }
            }
        }
        
        return ($Names | Sort-Object) -as [string[]]
    }
    catch {
        Write-Log -Message "Hiba az XML olvasásakor: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

function Get-OnkormanyzatAttendanceData {
    try {
        [xml]$xml = Get-Content $Script:AttendanceXmlPath
    }
    catch {
        Write-Log -Message "Hiba az XML betöltésekor: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }

    $result = @()

    # Minden department, ahol van Önkormányzat pozíció
    $Departments = $xml.Jelenlet.Department |
                   Where-Object { $_.Position.Name -contains "Önkormányzat" }

    foreach ($Dept in $Departments) {

        $Facility = $Dept.Facility

        $Names = $Dept.Position |
                 Where-Object { $_.Name -eq "Önkormányzat" } |
                 ForEach-Object { $_.Nev }

        $result += [pscustomobject]@{
            Facility = $Facility
            Names    = $Names
        }
    }

    return $result
}



# Function to get facility name from a department
function Get-DepartmentFacility {
    param (
        [string]$DepartmentName
    )
    
    try {
        $XDoc = [System.Xml.Linq.XDocument]::Load($Script:AttendanceXmlPath)
        $Root = $XDoc.Root
        
        $Dept = $Root.Elements("Department") | Where-Object { $_.Attribute("Name").Value -eq $DepartmentName } | Select-Object -First 1
        if ($Dept) {
            return $Dept.Attribute("Facility").Value
        }
        
        return "Kecskeméti Fürdő"
    }
    catch {
        Write-Log -Message "Hiba az XML olvasásakor: $($_.Exception.Message)" -Level "ERROR"
        return "Kecskeméti Fürdő"
    }
}

# Function to open a file dialog
function Open-File([string] $initialDirectory) {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $initialDirectory
    $OpenFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.FileName
}

# Function to configure cell formatting
function Configure-CellFormatting {
    param (
        [object]$Cell,
        [string]$FontName = "Times New Roman",
        [int]$FontSize = 14,
        [bool]$Bold = $false
    )
    $Cell.Font.Name = $FontName
    $Cell.Font.Size = $FontSize
    $Cell.Font.Bold = $Bold
}

# Function to print a worksheet with data from a names array
function Print-Worksheet {
    param (
        [object]$Worksheet,
        [string]$HeaderText,
        [string[]]$Names,
        [int]$HeaderRow = 1,
        [int]$HeaderColumn = 9,
        [int]$DataStartRow = 1,
        [int]$DataStartColumn = 1
    )

    # Set header text
    $Worksheet.Cells.Item($HeaderRow, $HeaderColumn) = $HeaderText
    Configure-CellFormatting -Cell $Worksheet.Cells.Item($HeaderRow, $HeaderColumn)

    # Populate and print data. Each person is handled in its own try/catch so
    # a single bad entry (PrintOut failure, etc.) is logged and
    # skipped instead of aborting everyone still left in the list.
    $FailedCount = 0

    foreach ($Name in $Names) {
        try {
            $Worksheet.Cells.Item($DataStartRow, $DataStartColumn) = $Name
            Configure-CellFormatting -Cell $Worksheet.Cells.Item($DataStartRow, $DataStartColumn)
            $Worksheet.PrintOut()
        }
        catch {
            $FailedCount++
            Write-Log -Message "Nyomtatási hiba a név: '$Name': $($_.Exception.Message)" -Level "ERROR"
            # Deliberately continue - one bad row should not stop the rest of the list
        }
    }

    if ($FailedCount -gt 0) {
        Write-Log -Message "Figyelmeztetés: $FailedCount / $($Names.Count) bejegyzés nyomtatása sikertelen volt." -Level "WARNING"
    }
}

# Shared worker: does everything that used to be copy-pasted five times over -
# sets duplex mode, opens Excel/the workbook, prints the sheet, and cleans up
# the COM objects no matter what happens in between.
function Invoke-AttendanceSheetPrint {
    param (
        [Parameter(Mandatory = $true)][string]$OpenFile,
        [Parameter(Mandatory = $true)][string[]]$Names,
        [string]$SheetName = "Fizikai",
        [string]$HeaderText = "Kecskeméti Fürdő"
    )

    try {
        Set-DuplexingMode -Mode "Duplex"
    }
    catch {
        # By design: if we can't guarantee duplex mode, we must not print anything.
        Write-Log -Message "Hiba: Duplex mód nem állítható be, nyomtatás leállítva: $($_.Exception.Message)" -Level "ERROR"
        exit 1
    }

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Workbook = $null

    try {
        try {
            $Workbook = $Excel.Workbooks.Open($OpenFile)
        }
        catch {
            Write-Log -Message "Hiba: Munkafüzet nem nyitható meg: '$OpenFile': $($_.Exception.Message)" -Level "ERROR"
            exit 1
        }

        try {
            Print-Worksheet -Worksheet $Workbook.Sheets.Item($SheetName) `
                -HeaderText $HeaderText `
                -Names $Names
        }
        catch {
            Write-Log -Message "Hiba: Munkalap nyomtatása sikertelen: '$SheetName': $($_.Exception.Message)" -Level "ERROR"
            exit 1
        }
        finally {
            if ($Workbook) {
                try {
                    $Workbook.Close($false)
                }
                catch {
                    # Logged, not rethrown - we don't want a Close() failure to
                    # mask whatever the real error was above.
                    Write-Log -Message "Hiba: Munkafüzet bezárása sikertelen: '$OpenFile': $($_.Exception.Message)" -Level "ERROR"
                }
            }
        }
    }
    finally {
        try {
            $Excel.Quit()
        }
        catch {
            Write-Log -Message "Hiba: Excel kilépése sikertelen: $($_.Exception.Message)" -Level "ERROR"
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
}

function Get-SheetPath {
    $OpenFile = Open-File $env:USERPROFILE
    if (-not $OpenFile) {
        Write-Log -Message "Nem kiválasztott fájl. Kilépés." -Level "ERROR"
        exit
    }
    else {
        return $OpenFile
    }
}

# Each of these functions now gets names from the XML instead of CSV files
# and uses the facility name from the XML

function Print-AttandanceSheetUszomester {
    param (
        [string]$OpenFile
    )
    $Names = Get-DepartmentNames -DepartmentName "Fürdő"
    $Facility = Get-DepartmentFacility -DepartmentName "Fürdő"
    Invoke-AttendanceSheetPrint -OpenFile $OpenFile -Names $Names -HeaderText $Facility
}

function Print-AttandanceSheetFrontOffice {
    param (
        [string]$OpenFile
    )
    $Names = Get-DepartmentNames -DepartmentName "Front office"
    $Facility = Get-DepartmentFacility -DepartmentName "Front office"
    Invoke-AttendanceSheetPrint -OpenFile $OpenFile -Names $Names -HeaderText $Facility
}

function Print-AttandanceSheetGyogyaszat {
    param (
        [string]$OpenFile
    )
    $Names = Get-DepartmentNames -DepartmentName "Gyógyászat"
    $Facility = Get-DepartmentFacility -DepartmentName "Gyógyászat"
    Invoke-AttendanceSheetPrint -OpenFile $OpenFile -Names $Names -HeaderText $Facility
}

function Print-AttandanceSheetKarbantarto {
    param (
        [string]$OpenFile
    )
    $Names = Get-DepartmentNames -DepartmentName "Karbantartó"
    $Facility = Get-DepartmentFacility -DepartmentName "Karbantartó"
    Invoke-AttendanceSheetPrint -OpenFile $OpenFile -Names $Names -HeaderText $Facility
}

function Print-AttandanceSheetGepesz {
    param (
        [string]$OpenFile
    )
    $Names = Get-DepartmentNames -DepartmentName "Gépészet"
    $Facility = Get-DepartmentFacility -DepartmentName "Gépészet"
    Invoke-AttendanceSheetPrint -OpenFile $OpenFile -Names $Names -HeaderText $Facility
}

function Print-AttandanceSheetOnkormanyzat {
    param (
        [string]$OpenFile
    )

    $OnkormanyzatData = Get-OnkormanyzatAttendanceData

    foreach ($item in $OnkormanyzatData) {
        Invoke-AttendanceSheetPrint -OpenFile $OpenFile -Names $item.Names -HeaderText $item.Facility
    }
}

