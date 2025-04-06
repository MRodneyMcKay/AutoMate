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
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'TestSchoolDay\TestSchoolDay.psd1')

# Function to calculate the date of the next Monday
function Get-NextMonday {
    for ($i = 1; $i -le 7; $i++) {
        if ((Get-Date).AddDays($i).DayOfWeek -eq 'Monday') {
            return (Get-Date).AddDays($i)
        }
    }
}

# Function to handle document operations (open, replace, save, print)
function Process-Document {
    param (
        [string]$Filename,
        [datetime]$FindDate,
        [datetime]$ReplaceDate,
        [ref]$WordApp
    )

    # Open document
    try {
        $Document = $WordApp.Value.Documents.Open($Filename)
        if ($Document) {
            Write-Log -Message "Opened document: $Filename"
        } else {
            throw
        }
    }
    catch {
        Write-Log -Message "Failed to open document: $Filename" -Level "ERROR"
        return $false
    }       

    # Set find and replace parameters
    $FindText = $FindDate.ToString("MMMM dd.")
    $ReplaceText = $ReplaceDate.ToString("MMMM dd.")
    $MatchCase = $false
    $MatchWholeWord = $true
    $MatchWildcards = $false
    $Forward = $false
    $Wrap = 1
    $Replace = 2

    # Perform find and replace
    if (!$Document.Content.Find.Execute($FindText, $MatchCase, $MatchWholeWord, $MatchWildcards, $null, $null, $Forward, $Wrap, $null, $ReplaceText, $Replace)) {
        Write-Log -Message "Text not found: $FindText" -Level "ERROR"
        $Document.Close(-1)
        return $false
    } else {
        Write-Log -Message "Replaced '$FindText' with '$ReplaceText' in document: $Filename"
    }

    # Save the document
    $Document.Save()
    Write-Log -Message "Saved document: $Filename"

    # Handle holiday-specific printing logic
    if (Test-Schoolday -date $ReplaceDate) {
        $Document.PrintOut()
        Write-Log -Message "Printed document: $Filename"
    }

    # Close the document
    $Document.Close(-1)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
    return $true
}

# Main script logic
$nextWeek = Get-NextMonday
$Word = New-Object -ComObject Word.Application
$basePath = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Úszó Nemzet Program"

for ($i = 1; $i -le 5; $i++) {
    # Generate day name and file path
    $dayName = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetDayName($nextWeek.DayOfWeek))
    $Filename = "$($basePath)\-=SABLON=-\$dayName.docx"

    # Ensure file is writable
    Set-ItemProperty -Path $Filename -Name IsReadOnly -Value $false

    # Process the document
    if (-not (Process-Document -Filename $Filename -FindDate $nextWeek.AddDays(-7) -ReplaceDate $nextWeek -WordApp ([ref]$Word))) {
        $nextWeek = $nextWeek.AddDays(1)
        continue
    }

    # Mark file as read-only
    Set-ItemProperty -Path $Filename -Name IsReadOnly -Value $true
    $nextWeek = $nextWeek.AddDays(1)
}

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) | Out-Null

# Perform garbage collection to finalize cleanup
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

<# SUMMERTIME
$today = Get-Date
$honap = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName((get-date).Month))
$base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Úszó Nemzet Program\'
$Filename = "$($base)-=SABLON=-\nyár.docx"
$SaveAsDir="$($base)$($today.Year).$($today.Month) - $honap\"
$SaveAs="$($SaveAsDir)$($today.ToString('yyyy.MM.dd.')).docx"
$Word=NEW-Object –comobject Word.Application
Set-ItemProperty -Path $Filename -Name IsReadOnly -Value $true
$Document=$Word.documents.open($Filename)
$FindText = "Nap"
$ReplaceText = $today.ToString("MMMM dd.") # <= Replace it with this text
$MatchCase = $false
$MatchWholeWorld = $true
$MatchWildcards = $false
$MatchSoundsLike = $false
$MatchAllWordForms = $false
$Forward = $false
$Wrap = 1
$Format = $false
$Replace = 2
if (!$($Document.Content.Find.Execute($FindText, $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $ReplaceText, $Replace))) {
    $Document.Close(-1)
    $Word.Quit()
}
else {   

    if (-Not ([System.IO.Directory]::Exists($SaveAsDir)))
    {
        mkdir -p $SaveAsDir
    }
    $Document.SaveAs2($SaveAs)
    $Word.visible=$true
}
#>
