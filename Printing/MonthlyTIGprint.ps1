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

# Load the required Word Interop assembly
function Load-WordInterop {
    Add-Type -Path "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
}

# Get the default printer
function Get-DefaultPrinter {
    return (Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true })
}

# Perform the mail merge
function Perform-MailMerge {
    param (
        [string]$DocumentPath,
        [string]$DataSourcePath
    )

    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false

    $doc = $Word.Documents.Open($DocumentPath)
    $mailMerge = $doc.MailMerge
    $mailMerge.OpenDataSource($DataSourcePath)
    $mailMerge.Destination = [Microsoft.Office.Interop.Word.WdMailMergeDestination]::wdSendToNewDocument
    $mailMerge.Execute($true)

    return @{
        Word = $Word
        Document = $Word.ActiveDocument
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
        [string]$PrinterName,
        [string]$PageRange
    )

    # Configure printer for one-sided printing
    Set-PrintConfiguration -PrinterName $PrinterName -DuplexingMode OneSided | Out-Null

    # Print the document
    $BackGround = 0
    $Range = [Microsoft.Office.Interop.Word.WdPrintOutRange]::wdPrintRangeOfPages
    $Item = 0
    $Copies = 1
    $Document.PrintOut([ref]$BackGround, [Type]::Missing, [ref]$Range, [Type]::Missing, [Type]::Missing, [Type]::Missing, [ref]$Item, [ref]$Copies, [ref]$PageRange, [Type]::Missing)

    # Restore printer to two-sided printing
    Set-PrintConfiguration -PrinterName $PrinterName -DuplexingMode TwoSidedLongEdge
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
}

# Main script logic
Load-WordInterop

$defaultPrinter = Get-DefaultPrinter
$documentPath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\TIG jelenléti ív_2025.docx"
$dataSourcePath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\TIG csoportok.csv"

$mailMergeResult = Perform-MailMerge -DocumentPath $documentPath -DataSourcePath $dataSourcePath
$mergedDocument = $mailMergeResult.Document
$Word = $mailMergeResult.Word
$OriginalDocument = $mailMergeResult.OriginalDocument

$month = (Get-Date).Month + 1
$sections = $mergedDocument.Sections.Count
$pageRange = Generate-PageRange -Month $month -Sections $sections

Print-Document -Document $mergedDocument -PrinterName $defaultPrinter.Name -PageRange $pageRange

Cleanup-WordObjects -Word $Word -Document $mergedDocument -OriginalDocument $OriginalDocument