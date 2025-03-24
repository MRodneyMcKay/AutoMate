Import-Module "$PSScriptRoot\TestSchoolday.psm1"

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
    $Document = $WordApp.Value.Documents.Open($Filename)

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
        Write-Output "Text not found: $FindText"
        $Document.Close(-1)
        return $false
    }

    # Save the document
    $Document.Save()

    # Handle holiday-specific printing logic
    if (Test-Schoolday -date $ReplaceDate) {
        $Document.PrintOut()
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
