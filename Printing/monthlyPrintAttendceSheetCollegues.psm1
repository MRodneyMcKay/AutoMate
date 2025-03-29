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

# Function to print a worksheet with data from a CSV file
function Print-Worksheet {
    param (
        [object]$Worksheet,
        [string]$HeaderText,
        [string]$CsvPath,
        [int]$HeaderRow = 1,
        [int]$HeaderColumn = 9,
        [int]$DataStartRow = 1,
        [int]$DataStartColumn = 1
    )

    # Set header text
    $Worksheet.Cells.Item($HeaderRow, $HeaderColumn) = $HeaderText
    Configure-CellFormatting -Cell $Worksheet.Cells.Item($HeaderRow, $HeaderColumn)

    # Populate and print data
    Import-CSV -Path $CsvPath | Sort-Object Név | ForEach-Object {
        $Worksheet.Cells.Item($DataStartRow, $DataStartColumn) = $_.Név
        Configure-CellFormatting -Cell $Worksheet.Cells.Item($DataStartRow, $DataStartColumn)
        $Worksheet.PrintOut()
    }
}

# Main script logic


function Get-SheetPath {
    $OpenFile = Open-File $env:USERPROFILE
    if (-not $OpenFile) {
        Write-Output "No file selected. Exiting."
        exit
    }
    else {
        return $OpenFile
    }
}

function Print-AttandanceSheetUszomester {
    param (
        [string]$OpenFile
    )

    # Get the default printer
    $DefaultPrinter = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }
    Set-PrintConfiguration -PrinterName $DefaultPrinter.Name -DuplexingMode TwoSidedLongEdge
    
    # Open Excel workbook
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Workbook = $Excel.Workbooks.Open($OpenFile)
    Print-Worksheet -Worksheet $Workbook.Sheets.Item("Fürdő") `
    -HeaderText "Kecskeméti Fürdő" `
    -CsvPath "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\Fürdő.csv"

    # Close workbook and clean up
    $Workbook.Close($false)
    $Excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

function Print-AttandanceSheetFrontOffice {
    param (
        [string]$OpenFile
    )

    # Get the default printer
    $DefaultPrinter = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }
    Set-PrintConfiguration -PrinterName $DefaultPrinter.Name -DuplexingMode TwoSidedLongEdge
    
    # Open Excel workbook
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Workbook = $Excel.Workbooks.Open($OpenFile)
    Print-Worksheet -Worksheet $Workbook.Sheets.Item("Jegypénztár, recepció") `
    -HeaderText "Kecskeméti Fürdő" `
    -CsvPath "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\Jegypénztár, recepció.csv"

    # Close workbook and clean up
    $Workbook.Close($false)
    $Excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

function Print-AttandanceSheetKarbantarto {
    param (
        [string]$OpenFile
    )

    # Get the default printer
    $DefaultPrinter = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }
    Set-PrintConfiguration -PrinterName $DefaultPrinter.Name -DuplexingMode TwoSidedLongEdge
    
    # Open Excel workbook
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Workbook = $Excel.Workbooks.Open($OpenFile)

    Print-Worksheet -Worksheet $Workbook.Sheets.Item("Karbantartó") `
    -HeaderText "Kecskeméti Fürdő" `
    -CsvPath "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\Karbantartó.csv" `
    -HeaderRow 2 -HeaderColumn 11 -DataStartRow 2 -DataStartColumn 2


    # Close workbook and clean up
    $Workbook.Close($false)
    $Excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

function Print-AttandanceSheetGepesz {
    param (
        [string]$OpenFile
    )

    # Get the default printer
    $DefaultPrinter = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }
    Set-PrintConfiguration -PrinterName $DefaultPrinter.Name -DuplexingMode TwoSidedLongEdge
    
    # Open Excel workbook
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Workbook = $Excel.Workbooks.Open($OpenFile)

    $sheetToCopy = $workbook.Sheets.Item("Karbantartó") # Replace with your sheet name
    $sheetToCopy.Copy($workbook.Sheets.Item($workbook.Sheets.Count)) # Copy to the end of the workbook
    
    # Rename the copied worksheet
    $copiedSheet = $workbook.Sheets.Item($sheetToCopy.Index + 1) # Access the newly created sheet
    $copiedSheet.Name = "Gépész" # Replace with your desired sheet name
    $copiedSheet.Cells.Item(1,2) = $copiedSheet.Cells.Item(1,2).Text.Replace('KARBANTARTÓ', 'GÉPÉSZ')
    Print-Worksheet -Worksheet $Workbook.Sheets.Item("Gépész") `
        -HeaderText "Kecskeméti Fürdő" `
        -CsvPath "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\Gépész.csv" `
        -HeaderRow 2 -HeaderColumn 11 -DataStartRow 2 -DataStartColumn 2

    # Close workbook and clean up
    $Workbook.Close($false)
    $Excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)

    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}