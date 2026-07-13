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

$global:DataFolder = "C:\Users\Hirossport\Hiros Sport Nonprofit Kft\Hiros-sport - Dokumentumok\Furdo\Recepcio\Nyomtatni\Jelenlétik, igények"
$global:XmlPath = Join-Path -Path $global:DataFolder -ChildPath "nevek.xml"
$global:CultureHU = [System.Globalization.CultureInfo]::GetCultureInfo("hu-HU")

$global:Data = [ordered]@{}
$global:DeptControls = @{}
$global:DirtyDepartments = New-Object System.Collections.Generic.HashSet[string]
$global:DirtyPositions = New-Object System.Collections.Generic.HashSet[string]

function Initialize-EditorState {
    $global:Data = [ordered]@{}
    $global:DeptControls = @{}
    $global:DirtyDepartments = New-Object System.Collections.Generic.HashSet[string]
    $global:DirtyPositions = New-Object System.Collections.Generic.HashSet[string]
}
