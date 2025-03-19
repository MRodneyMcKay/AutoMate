# Function to configure the printer
function Configure-Printer {
    param (
        [string]$PrinterName,
        [string]$DuplexingMode
    )
    Set-PrintConfiguration -PrinterName $PrinterName -DuplexingMode $DuplexingMode | Out-Null
}

# Function to open a Word document and perform mail merge
function Perform-MailMerge {
    param (
        [string]$DocumentPath,
        [string]$DataSourcePath,
        [string]$ConnectionString,
        [string]$SQLStatement
    )

    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false

    $Document = $Word.Documents.Open($DocumentPath)
    $MailMerge = $Document.MailMerge
    $MailMerge.OpenDataSource(
        $DataSourcePath,
        [ref][Microsoft.Office.Interop.Word.WdOpenFormat]::wdOpenFormatAuto,
        [ref]$false, # ConfirmConversions
        [ref]$true,  # ReadOnly
        [ref]$true,  # LinkToSource
        [ref]$false, # AddToRecentFiles
        [Type]::Missing, # PasswordDocument
        [Type]::Missing, # PasswordTemplate
        [ref]$false, # Revert
        [Type]::Missing, # WritePasswordDocument
        [Type]::Missing, # WritePasswordTemplate
        [ref]$ConnectionString, # Connection
        [ref]$SQLStatement, # SQLStatement
        [Type]::Missing # SQLStatement1
    )

    $MailMerge.Destination = [Microsoft.Office.Interop.Word.WdMailMergeDestination]::wdSendToNewDocument
    $MailMerge.Execute($true)

    return @{
        Word = $Word
        MergedDocument = $Word.ActiveDocument
        OriginalDocument = $Document
    }
}

# Function to generate the page range
function Generate-PageRange {
    param (
        [int]$Month,
        [int]$Sections
    )

    $Pages = ""
    for ($s = 1; $s -lt $Sections; $s++) {
        $Pages += "p$($Month)s$s;"
    }
    $Pages += "p$($Month)s$Sections"
    return $Pages
}

# Function to print the document
function Print-Document {
    param (
        [object]$Document,
        [string]$PrinterName,
        [string]$PageRange
    )

    Configure-Printer -PrinterName $PrinterName -DuplexingMode "OneSided"

    $Background = 0
    $Range = [Microsoft.Office.Interop.Word.WdPrintOutRange]::wdPrintRangeOfPages
    $Item = 0
    $Copies = 1

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

    Configure-Printer -PrinterName $PrinterName -DuplexingMode "TwoSidedLongEdge"
}

# Function to clean up Word COM objects
function Cleanup-WordObjects {
    param (
        [object]$Word,
        [object]$MergedDocument,
        [object]$OriginalDocument
    )

    $OriginalDocument.Close($false)
    $MergedDocument.Close($false)
    $Word.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OriginalDocument) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MergedDocument) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
}

function Print-CommutingAllowance {
# Main script logic
$bejaroPath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Útiköltség nyomtatvány_2025.docx"
$bejarodataSourcePath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Útiköltség nyomtatvány.xlsx"
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$bejarodataSourcePath;Mode=Read;Extended Properties=`"HDR=YES;IMEX=1;`";"
$sqlStatement = "SELECT * FROM [Munka1$]"

$defaultPrinter = (Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }).Name

# Perform mail merge
$mailMergeResult = Perform-MailMerge -DocumentPath $bejaroPath -DataSourcePath $bejarodataSourcePath -ConnectionString $connectionString -SQLStatement $sqlStatement
$mergedDocument = $mailMergeResult.MergedDocument
$wordApp = $mailMergeResult.Word
$originalDocument = $mailMergeResult.OriginalDocument

# Generate page range and print
$month = (Get-Date).Month + 1
$sections = $mergedDocument.Sections.Count
$pageRange = Generate-PageRange -Month $month -Sections $sections

Print-Document -Document $mergedDocument -PrinterName $defaultPrinter -PageRange $pageRange

# Clean up
Cleanup-WordObjects -Word $wordApp -MergedDocument $mergedDocument -OriginalDocument $originalDocument
}

