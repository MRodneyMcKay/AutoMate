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

[System.Reflection.Assembly]::LoadFrom([System.Environment]::GetEnvironmentVariable("OfficeAssemblies_Word", [System.EnvironmentVariableTarget]::User)) 
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath '..\Printing\TestSchoolday.psm1')
# Get the current day of the week
$today = (Get-Date)

# Check if today is Saturday or Sunday
if ($today.DayOfWeek -in @('Saturday', 'Sunday')) {
    exit
}

if (Test-Schoolday) {
    $nap = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).DateTimeFormat.GetDayName($today.DayOfWeek))
    $base='C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Úszó Nemzet Program\'
    $SaveAsDir="$($base)$($today.ToString('yyyy.MM')) - $((Get-Culture).DateTimeFormat.GetMonthName($today.Month))\"
    $SaveAs="$($SaveAsDir)$($today.ToString('yyyy.MM.dd')).docx"
    $Filename = "$($base)-=SABLON=-\$nap.docx"
    $Word=NEW-Object –comobject Word.Application
    $Word.visible=$true
    if (-Not (Test-Path $SaveAsDir)) {
        New-Item -Path $SaveAsDir -ItemType Directory -Force
    }
    if (-Not (Test-Path $SaveAs)) {
        $Document=$Word.documents.open($Filename)
        $Document.SaveAs2($SaveAs)
    } else{
        $Document=$Word.documents.open($SaveAs)
    }
}