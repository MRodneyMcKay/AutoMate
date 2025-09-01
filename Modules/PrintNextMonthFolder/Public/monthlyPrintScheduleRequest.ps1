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

# Function to open a Word document
function Open-WordDocument {
    param (
        [string]$FilePath
    )
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    try {
        $Document = $Word.Documents.Open($FilePath)
        if ($Document) {
            Write-Log -Message "Document opened successfully: $FilePath"
        } else {
            throw
        }

    } catch {
        Write-log "Failed to open document: $FilePath" -Level "ERROR"
        return $null
    }
    
    return @{
        Word = $Word
        Document = $Document
    }
}

# Function to print a specific page range
function Print-DocumentSchedule {
    param (
        [object]$Document,
        [string]$PageRange,
        [int]$Copies = 1
    )
    $Background = 0
    $Range = [Microsoft.Office.Interop.Word.WdPrintOutRange]::wdPrintRangeOfPages
    $Item = 0
    Set-PrinterDuplexMode -DuplexingMode "TwoSidedLongEdge"
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
     Write-Log -Message "Printing $PageRange pages of $($Document.Name)"
}

# Function to clean up Word COM objects
function Cleanup-WordObjects_requests {
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
    $Start = (($MonthOffset + (Get-Date).Month) * 6) + $StartMultiplier
    $End = (($MonthOffset + (Get-Date).Month) * 6) + $EndMultiplier
    return "$Start-$End"
}

# Main function to handle printing
function Print-Igeny {
    param (
        [string]$FilePath,
        [int]$StartMultiplier,
        [int]$EndMultiplier
    )

    $WordObjects = Open-WordDocument -FilePath $FilePath
    $Word = $WordObjects.Word
    $Document = $WordObjects.Document

    $PageRange = Get-PageRange -MonthOffset 1 -StartMultiplier $StartMultiplier -EndMultiplier $EndMultiplier
    Print-DocumentSchedule -Document $Document -PageRange $PageRange

    Cleanup-WordObjects_requests -Word $Word -Document $Document
}

# Define the file path


function Print-RequestUszomester {
    param (
        [string]$FilePath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\Igények.docx"
    )
    Print-Igeny -FilePath $FilePath -StartMultiplier 1 -EndMultiplier 2
}

function Print-RequestFrontOffice {
    param (
        [string]$FilePath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\Igények.docx"
    )
    Print-Igeny -FilePath $FilePath -StartMultiplier 3 -EndMultiplier 4
}

function Print-RequestGyogyaszat {
    param (
        [string]$FilePath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények\Igények.docx"
    )
    Print-Igeny -FilePath $FilePath -StartMultiplier 5 -EndMultiplier 6
}