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

# --- Global country mappings ---
$ovalToIso = @{
    "H"="HU"; "D"="DE"; "A"="AT"; "F"="FR"; "I"="IT"; "CH"="CH"; "NL"="NL"; "B"="BE"; "CZ"="CZ"; "SK"="SK"
    "PL"="PL"; "GB"="GB"; "E"="ES"; "S"="SE"; "FIN"="FI"; "EST"="EE"; "LT"="LT"; "LV"="LV"
    "RO"="RO"; "BG"="BG"; "HR"="HR"; "GR"="GR"; "DK"="DK"; "IRL"="IE"; "L"="LU"; "P"="PT"
    "M"="MT"; "CY"="CY"; "SI"="SI"; "TR"="TR"; "USA"="US"; "CDN"="CA"; "J"="JP"; "CN"="CN"
    "KR"="KR"; "IND"="IN"; "UA"="UA"; "RUS"="RU"; "BR"="BR"; "MEX"="MX"; "C"="CU"
}

$euCountries = @("AT","BE","BG","HR","CY","CZ","DK","EE","FI","FR","DE","GR","HU","IE","IT",
                 "LV","LT","LU","MT","NL","PL","PT","RO","SK","SI","ES","SE")

$allowedNonEU = @("CN", "RU", "JP", "IN", "TR", "UA", "US", "CA", "BR", "MX", "GB")

$isoToHunName = @{
    "HU" = "Magyarország"; "DE" = "Németország"; "FR" = "Franciaország"; "AT" = "Ausztria"
    "IT" = "Olaszország"; "SK" = "Szlovákia"; "CZ" = "Csehország"; "RO" = "Románia"
    "PL" = "Lengyelország"; "SI" = "Szlovénia"; "HR" = "Horvátország"; "BG" = "Bulgária"
    "ES" = "Spanyolország"; "PT" = "Portugália"; "NL" = "Hollandia"; "BE" = "Belgium"
    "LU" = "Luxemburg"; "IE" = "Írország"; "FI" = "Finnország"; "SE" = "Svédország"
    "DK" = "Dánia"; "LT" = "Litvánia"; "LV" = "Lettország"; "EE" = "Észtország"
    "GR" = "Görögország"; "CY" = "Ciprus"; "MT" = "Málta"; "TR" = "Törökország"
    "IL" = "Izrael"; "UA" = "Ukrajna"; "RU" = "Oroszország"; "CH" = "Svájc"
    "GB" = "Egyesült Királyság"; "US" = "Amerikai Egyesült Államok"; "CA" = "Kanada"
    "CN" = "Kína"; "JP" = "Japán"; "KR" = "Dél-Korea"; "IN" = "India"; "BR" = "Brazília"
    "MX" = "Mexikó"; "CU" = "Kuba"
}

# --- Normalization function ---
function Normalize-NationalityCode {
    param($gyujtendo, $megnevezes)

    $code = $gyujtendo.Trim().ToUpper()
    if ($ovalToIso.ContainsKey($code)) {
        $code = $ovalToIso[$code]
    }

    if ([string]::IsNullOrWhiteSpace($megnevezes)) {
        return "HU"
    }

    if ($code -eq "HU" -or $code -in $euCountries -or $code -in $allowedNonEU) {
        return $code
    }

    return "HU"
}