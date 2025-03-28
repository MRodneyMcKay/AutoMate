<#  
    This file is part of sturdy-octo-garbanzo.  

    sturdy-octo-garbanzo is free software: you can redistribute it and/or modify  
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

function create-Directories {
    param (
        [datetime]$Today = (Get-Date)
    )
    if ($Today.DayOfWeek.value__ -eq 0)
    {
        $Monday = $Today.AddDays(-6 - $Today.DayOfWeek.value__)
        $Sunday = $Monday.AddDays(6)
        $Week = $(get-date $Today -UFormat %V)
    }
    else {
        $Monday = $Today.AddDays(1 - $Today.DayOfWeek.value__)
        $Sunday = $Monday.AddDays(6)
        $Week = $(get-date $Today -UFormat %V)
    }
    #HIBA mappa
    $base="C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\HIBA\HIBA $($Today.Year)\"
    $dir = "$base$Week. hét ($($Monday.ToString("MM.dd.")) - $($Sunday.ToString("MM.dd.")))"
    if (-Not ([System.IO.Directory]::Exists($dir)))
    {
        mkdir -p $dir
    }   
}