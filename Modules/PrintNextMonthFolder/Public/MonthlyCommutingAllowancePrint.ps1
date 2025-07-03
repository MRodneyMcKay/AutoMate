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
function Print-CommutingAllowancePages {
    param (
        [object]$Document,
        [string]$PageRange
    )

    Set-PrinterDuplexMode -DuplexingMode "OneSided"

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

    Set-PrinterDuplexMode -DuplexingMode "TwoSidedLongEdge"
}

# Function to clean up Word COM objects
function Cleanup-WordObjects_Allowance {
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

Print-CommutingAllowancePages -Document $mergedDocument -PrinterName $defaultPrinter -PageRange $pageRange

Write-Log -Message "Printing $pageRange for $($mergedDocument.Name) on $defaultPrinter"

# Clean up
Cleanup-WordObjects_Allowance -Word $wordApp -MergedDocument $mergedDocument -OriginalDocument $originalDocument
}

