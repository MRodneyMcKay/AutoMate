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

function Get-CleanedUpZipCodes {
    param(
        # Path to the CSV file encoded as Windows-1250
        [Parameter(Mandatory = $true)]
        [string]$csvPath,

        # The desired total number of visitors
        [Parameter(Mandatory = $true)]
        [int]$DesiredTotal
    )

    if ($null -eq $csvPath) {
        throw "CSV path is null. Path: $csvPath"
    }

    # Import the CSV file using the specified encoding.
    # This CSV is assumed to have headers: "Gyűjtendő", "Megnevezés", "Darabszám"
    $csvData = Import-Csv -Path $CsvPath -Encoding "Windows-1250" -Delimiter ";"

    # Initialize a hashtable that will aggregate visitor numbers per zip code.
    $aggregated = @{}

    foreach ($row in $csvData) {
        # Skip the summary row (the CSV ends with "Mindösszesen")
        if ($row.'Gyűjtendő' -eq "Mindösszesen") {
            continue
        }

        $zip = $row.'Gyűjtendő'
        $visitorCount = 0
        [int]::TryParse($row.'Darabszám', [ref] $visitorCount) | Out-Null

        # Check if the zip code is exactly 4 digits.
        if ($zip -match '^\d{4}$') {
            $finalZip = $zip
        }
        else {
            # If not, use "6000" as the key.
            $finalZip = '6000'
        }

        # Accumulate the visitor count for the zip code.
        if ($aggregated.ContainsKey($finalZip)) {
            $aggregated[$finalZip] += $visitorCount
        }
        else {
            $aggregated[$finalZip] = $visitorCount
        }
    }

    # Calculate the current total of visitors from all the aggregated zip codes.
    $currentTotal = ($aggregated.Values | Measure-Object -Sum).Sum

    # Determine the adjustment needed so that the total equals the DesiredTotal.
    $difference = $DesiredTotal - $currentTotal

    # Adjust the "6000" entry accordingly.
    if ($aggregated.ContainsKey('6000')) {
        $aggregated['6000'] += $difference
    }
    else {
        # If no row was assigned to "6000", create one with the correction amount.
        $aggregated['6000'] = $difference
    }

    # Prepare the output as an array of custom objects.
    $resultArray = @()
    foreach ($zip in $aggregated.Keys | Sort-Object) {
        $resultArray += [PSCustomObject]@{
            Zip         = $zip
            Megnevezes  = if ($zipToName.ContainsKey($zip)) { $zipToName[$zip] } else { "Ismeretlen" }
            Darabszám   = [int]$aggregated[$zip]
        }
    }

    # Append the final record summarizing the total visitors.
    $totalVisitors = ($aggregated.Values | Measure-Object -Sum).Sum
    $resultArray += [PSCustomObject]@{
        Zip         = "Összesen"
        Darabszám   = [int]$totalVisitors
    }

    return $resultArray | Sort-Object Darabszám
}

function Get-ZipStats {
    [CmdletBinding()]
    param(
        # The cleaned-up CSV data; an array of PSCustomObjects with properties:
        # Zip, Megnevezés, and Darabszám.
        [Parameter(Mandatory = $true)]
        [array]$CleanedData
    )

    # Define the zip codes that belong to "Kecskemét környéke"
    $kecskemetKornyekeZips = @(
        "6076","6035","6055","6042","6034","6078","6044","6008",
        "6041","6043","6045","6050","6065","6032","6077","6031",
        "6060","6062","6064","6033","6070"
    )

    # Get Kecskemét count from zip code "6000"
    $kecskemetRecord = $CleanedData | Where-Object { $_.Zip -eq "6000" }
    $kecskemetCount = if ($kecskemetRecord) { [int]$kecskemetRecord.Darabszám } else { 0 }

    # Get the sum for "Kecskemét környéke"
    $kornyekeCount = ($CleanedData | Where-Object { $kecskemetKornyekeZips -contains $_.Zip } |
                      Measure-Object -Property Darabszám -Sum).Sum
    if (-not $kornyekeCount) { $kornyekeCount = 0 }

    # Determine the overall total.
    # We assume the cleaned data includes a summary row (with Zip = "Total")
    $totalRecord = $CleanedData | Where-Object { $_.Zip -eq "Összesen" }
    if ($totalRecord) {
        $overallTotal = [int]$totalRecord.Darabszám
    }
    else {
        # If no summary row exists, calculate the total from all records
        $overallTotal = ($CleanedData | Measure-Object -Property Darabszám -Sum).Sum
    }

    # "Egyéb magyarországon belüli település" is the remainder
    $egyebCount = $overallTotal - ($kecskemetCount + $kornyekeCount)

    # Create an output array with the statistics
    $statsArray = @()
    $statsArray = [PSCustomObject]@{
        Kecskemet = [int]$kecskemetCount
        Kecskemet_kornyeke = [int]$kornyekeCount
        Egyeb_telepules= [int]$egyebCount
        Osszesen = [int]$overallTotal
    }

    return $statsArray
}