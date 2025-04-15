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

# --- Function to load and group data ---
function Get-GroupedNationalityData {
    param (
        $csvPath
    )

    if ($null -eq $csvPath) {
        throw "CSV path is null. Path: $csvPath"
    }

    $data = Import-Csv $csvPath -Encoding "Windows-1250" -Delimiter ";"
    $records = $data | Where-Object { $_."Gyűjtendő" -notmatch "mindösszes" }

    $normalized = $records | ForEach-Object {
        $normCode = Normalize-NationalityCode $_."Gyűjtendő" $_."Megnevezés"
        [PSCustomObject]@{
            ISO   = $normCode
            Darabszám = [int]::Parse($_.Darabszám.Replace(',',''), [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }

    $grouped = $grouped = New-Object System.Collections.Generic.List[PSCustomObject]

    $normalized | Group-Object ISO | ForEach-Object {
        $iso = $_.Name
        $name = if ($isoToHunName.ContainsKey($iso)) { $isoToHunName[$iso] } else { "Ismeretlen" }
    
        $grouped.Add([PSCustomObject]@{
            ISO    = $iso
            Orszag = $name
            Darabszám  = [int]($_.Group | Measure-Object Darabszám -Sum).Sum
        })
    }
  

    # Add summary row
    $total = ($grouped | Measure-Object Darabszám -Sum).Sum
    $grouped += [PSCustomObject]@{
        ISO    = "Összesen"
        Orszag = ""
        Darabszám  = [int]$total
    }

    $grouped = $grouped | Sort-Object Darabszám

    return $grouped
}

# --- Function to create summary statistics ---
function Get-NationalityStats {
    param (
        [array]$grouped
    )

    $belfoldi = [int]($grouped | Where-Object { $_.ISO -eq "HU" }).Darabszám
    $eu = [int]($grouped | Where-Object { $_.ISO -ne "HU" -and $_.ISO -in $euCountries } | Measure-Object -Property Darabszám -Sum).Sum
    $noneu = [int]($grouped | Where-Object {$_.ISO -ne "Összesen" -and $_.ISO -ne "HU" -and $_.ISO -notin $euCountries } | Measure-Object -Property Darabszám -Sum).Sum

    $stats = [PSCustomObject]@{
        Belfoldi_HU       = [int]$belfoldi
        Kulfold_EU        = [int]$eu
        Kulfold_NonEU     = [int]$noneu
        Kulfold_Osszesen  = [int]$eu + $noneu
    }

    return $stats
}
