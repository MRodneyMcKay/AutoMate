# Function to configure the printer
function Configure-Printer {
    param (
        [string]$PrinterName,
        [string]$DuplexingMode
    )
    #Set-PrintConfiguration -PrinterName $PrinterName -DuplexingMode $DuplexingMode | Out-Null
}

# Function to open a Word document
function Open-WordDocument {
    param (
        [string]$FilePath
    )
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Document = $Word.Documents.Open($FilePath)
    return @{
        Word = $Word
        Document = $Document
    }
}

# Function to print a specific page range
function Print-Document {
    param (
        [object]$Document,
        [string]$PageRange,
        [int]$Copies = 1
    )
    $Background = 0
    $Range = [Microsoft.Office.Interop.Word.WdPrintOutRange]::wdPrintRangeOfPages
    $Item = 0

    $Document.PrintOut(
        [ref]$Background,
        [Type]::Missing,
        [ref]$Range,
        [Type]::Missing,
        [Type]::Missing,
        [Type]::Missing,
        [ref]$Item,
        [ref]$Copies,
        [ref]$PageRange,
        [Type]::Missing
    )
}

# Function to clean up Word COM objects
function Cleanup-WordObjects {
    param (
        [object]$Word,
        [object]$Document
    )
    $Document.Close($false)
    $Word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
}

# Function to calculate the page range
function Get-PageRange {
    param (
        [int]$MonthOffset,
        [int]$StartMultiplier,
        [int]$EndMultiplier
    )
    $Start = (($MonthOffset + (Get-Date).Month) * 4) + $StartMultiplier
    $End = (($MonthOffset + (Get-Date).Month) * 4) + $EndMultiplier
    return "$Start-$End"
}

# Main function to handle printing
function Print-Igeny {
    param (
        [string]$PrinterName =  (Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }).Name,
        [string]$FilePath,
        [int]$StartMultiplier,
        [int]$EndMultiplier
    )
    Configure-Printer -PrinterName $PrinterName -DuplexingMode "TwoSidedLongEdge"

    $WordObjects = Open-WordDocument -FilePath $FilePath
    $Word = $WordObjects.Word
    $Document = $WordObjects.Document

    $PageRange = Get-PageRange -MonthOffset 0 -StartMultiplier $StartMultiplier -EndMultiplier $EndMultiplier
    Print-Document -Document $Document -PageRange $PageRange

    Cleanup-WordObjects -Word $Word -Document $Document
}

# Define the file path
$FilePath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\Igények.docx"

# Print for Uszomester
Print-Igeny -FilePath $FilePath -StartMultiplier 1 -EndMultiplier 2

# Print for Front Office
Print-Igeny -FilePath $FilePath -StartMultiplier 3 -EndMultiplier 4