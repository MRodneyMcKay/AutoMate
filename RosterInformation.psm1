# Load the Microsoft.Office.Interop.Excel assembly

$possiblePaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\Microsoft.Office.Interop.Excel.dll",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\Microsoft.Office.Interop.Excel.dll",
    "C:\Program Files\Microsoft Office\Office16\Microsoft.Office.Interop.Excel.dll",
    "C:\Program Files (x86)\Microsoft Office\Office16\Microsoft.Office.Interop.Excel.dll",
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\*\Microsoft.Office.Interop.Excel.dll"
)

$dllPath = $possiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($dllPath) {
    Add-Type -Path $dllPath
} else {
    Write-Warning "Could not find Microsoft.Office.Interop.Excel in PS7!"
}

function Get-Receptionists {
    param (
        [datetime]$now = $(Get-Date)
    )

    $today = Get-Date $now
    $year = $today.Year
    $monthName = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetMonthName($today.Month))
    $pattern = '^(' + [regex]::escape($year) + ').*(' + [regex]::escape($monthName) + '|' + (Get-Culture).DateTimeFormat.GetMonthName($today.Month) + ').*(azd|ront).*xlsx$'
    $items = Get-ChildItem $([System.Environment]::GetFolderPath("Desktop")) | Where-Object { $_.Name.Normalize() -match ($pattern) } | Sort-Object Name -Descending | Select-Object -First 1
    if ($items.Count -eq 0) {
        Write-Error "No matching files found."
        return
    }

    $filePath = $items.FullName
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($filePath)
    $worksheet = $workbook.Sheets.Item(1)

    # Initialize an empty array
    $roster = @()

    try {        
        $roster += Get-Shift $worksheet 'Konfár Nikolett' $today
        $roster += Get-Shift $worksheet 'Raduska Zsolt' $today
    } catch {
        Write-Error "An error occurred: $_"
    } finally {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$workbook)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$worksheet)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable excel -ErrorAction SilentlyContinue
    }

    return $roster
}

function Get-Shift {
    param (
        [object]$worksheet,
        [string]$name,
        [datetime]$date
    )
    $found = $worksheet.Cells.Find($name)
    if ($found) {
        return [PSCustomObject]@{
            Name  = $name
            Shift = $worksheet.Cells($found.Row, 3 + $date.Day).Text
        }
    } else {
        Write-Error "$name not found in the worksheet."
        return $null
    }
}