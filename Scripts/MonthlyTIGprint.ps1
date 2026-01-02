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

$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$resolvedModulegPath = (Resolve-Path -Path $ModulePath).Path
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'LoggingSystem\LoggingSystem.psd1')

# Get the default printer
function Get-DefaultPrinter {
    return (Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true })
}

function Set-PrinterDuplexMode {
    param (
        [string]$PrinterName = (Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }).Name,
        [string]$DuplexingMode
    )
    try {
        Set-PrintConfiguration -PrinterName $PrinterName -DuplexingMode $DuplexingMode | Out-Null
        Write-Log -Message "Configure printer $PrinterName with duplexing mode $DuplexingMode." -Level INFO
    } catch {
        Write-Log -Message "Failed to configure printer $PrinterName with duplexing mode $DuplexingMode. Error: $_" -Level Error
    }
}

# Perform the mail merge
function Perform-MailMerge {
    param (
        [string]$DocumentPath,
        [string]$DataSourcePath
    )

    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false

    try {
        $doc = $Word.Documents.Open($DocumentPath)
        if ($doc) {
            Write-Log -Message "Document opened successfully: $DocumentPath" -Level "INFO"
        } else {
            throw
        }
    } catch {
        Write-Log -Message "Failed to open document: $DocumentPath" -Level "ERROR"
        return
    }

    $mailMerge = $doc.MailMerge
    $mailMerge.OpenDataSource($DataSourcePath)
    $mailMerge.Destination = [Microsoft.Office.Interop.Word.WdMailMergeDestination]::wdSendToNewDocument
    $mailMerge.Execute($true)

    foreach ($tbl in  $Word.ActiveDocument.Tables) {
        $tableRange = $tbl.Range
        $start = [Math]::Max(0, $tableRange.Start - 500)
        $end = [Math]::Min( $Word.ActiveDocument.Content.End, $tableRange.End + 500)
        $contextRange =  $Word.ActiveDocument.Range($start, $end)
        $contextText = $contextRange.Text

        if ($contextText -like "*Kecskeméti Piarista Gimnázium*") {
            try {
                $tbl.Columns.Item(3).Select() # Select the 3rd column

                    # Insert a column before the selected column
                    $Word.Selection.InsertColumns() 

                    $tbl.Cell(1, 3).Range.Text = "Osztály"
                    Write-Log -Message "Inserted 'Osztály' column in table near Piarista text." -Level "INFO"
            } catch {
                Write-Log -Message "Failed to modify table near Piarista text. Error: $_" -Level "WARNING"
            }
        }
    }

    return @{
        Word             = $Word
        Document         = $Word.ActiveDocument
        OriginalDocument = $doc
    }
}


# Generate the page range string
function Generate-PageRange {
    param (
        [int]$Month,
        [int]$Sections
    )

    $pages = ""
    for ($s = 1; $s -lt $Sections; $s++) {
        $pages += "p$($Month)s$s;"
    }
    $pages += "p$($Month)s$Sections"
    return $pages
}

# Print the document
function Print-Document {
    param (
        [object]$Document,
        [string]$PageRange
    )

    # Configure printer for one-sided printing
    Set-PrinterDuplexMode -DuplexingMode "OneSided"

    # Print the document
    $BackGround = 0
    $Range = [Microsoft.Office.Interop.Word.WdPrintOutRange]::wdPrintRangeOfPages
    $Item = 0
    $Copies = 1
    $Document.PrintOut([ref]$BackGround, [Type]::Missing, [ref]$Range, [Type]::Missing, [Type]::Missing, [Type]::Missing, [ref]$Item, [ref]$Copies, [ref]$PageRange, [Type]::Missing)
    Write-Log -Message "Printing $PageRange pages of $($Document.Name)"

    # Restore printer to two-sided printing
    Set-PrinterDuplexMode -DuplexingMode "TwoSidedLongEdge"
}

# Clean up and release COM objects
function Cleanup-WordObjects {
    param (
        [object]$Word,
        [object]$Document,
        [object]$OriginalDocument
    )

    $OriginalDocument.Close($false)
    $Document.Close($false)
    $Word.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OriginalDocument) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null

    Remove-Item -Path "$env:USERPROFILE\Desktop\MergedDocument.docx" -ErrorAction SilentlyContinue

}

function Get-NextMonthAndYear {
    param( [datetime]$Date = (Get-Date) )

    if ($Date.Month -eq 12) {
        return @{
            Month = 1
            Year  = $Date.Year + 1
        }
    } else {
        return @{
            Month = $Date.Month + 1
            Year  = $Date.Year
        }
    }
}


# Main script logic
[System.Reflection.Assembly]::LoadFrom(
    [System.Environment]::GetEnvironmentVariable("OfficeAssemblies_Word", [System.EnvironmentVariableTarget]::User)
)

$next = Get-NextMonthAndYear
$month = $next.Month
$year  = $next.Year

$documentPath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\TIG jelenléti ív_$year.docx"
$dataSourcePath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\TIG csoportok.csv"

$mailMergeResult = Perform-MailMerge -DocumentPath $documentPath -DataSourcePath $dataSourcePath

$mergedDocument   = $mailMergeResult.Document
$Word             = $mailMergeResult.Word
$OriginalDocument = $mailMergeResult.OriginalDocument

$sections = $mergedDocument.Sections.Count
$pageRange = Generate-PageRange -Month $month -Sections $sections

Print-Document -Document $mergedDocument -PageRange $pageRange

Cleanup-WordObjects -Word $Word -Document $mergedDocument -OriginalDocument $OriginalDocument
