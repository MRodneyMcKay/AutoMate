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

using namespace System.Collections.ObjectModel

class AttendenceDepartment {
    [string] $Name
    [string] $Facility
    [ObservableCollection[AttendencePosition]] $Positions

    AttendenceDepartment() {
        $this.Positions = [ObservableCollection[AttendencePosition]]::new()
    }
}
class AttendencePosition {
    [string] $Name
    [ObservableCollection[string]] $Names

    AttendencePosition() {
        $this.Names = [ObservableCollection[string]]::new()
    }
}

class AttendenceModel {
    [ObservableCollection[AttendenceDepartment]] $Departments
    [string] $XmlPath

    AttendenceModel([string] $xmlPath){
        $this.XmlPath = $xmlPath
        $this.Departments = [ObservableCollection[AttendenceDepartment]]::new()
        $this.Load()
    }
    [void] Load(){
        if(-not (Test-Path $this.XmlPath)){
            throw "XML not found: $($this.XmlPath)"
        }

        [xml]$xml = Get-Content $this.XmlPath -Raw
        foreach($departmentNode in $xml.Jelenlet.Department){
            $department = [AttendenceDepartment]::new()
            $department.Name = [string]$departmentNode.Name
            $department.Facility = [string]$departmentNode.Facility
            foreach($positionNode in $departmentNode.Position){
                $position = [AttendencePosition]::new()
                $position.Name = [string]$positionNode.Name
                foreach($nameNode in $positionNode.Nev){
                    $position.Names.Add([string]$nameNode)
                }
                $department.Positions.Add( $position )
            }
            $this.Departments.Add( $department )
        }
    }
    [void] Save(){
        $settings = [System.Xml.XmlWriterSettings]::new()
        $settings.Indent = $true
        $settings.Encoding = [System.Text.Encoding]::UTF8
        $writer = [System.Xml.XmlWriter]::Create( $this.XmlPath, $settings )
        try {
            $writer.WriteStartDocument()
            $writer.WriteStartElement( "Jelenlet" )
            foreach($department in $this.Departments){
                $writer.WriteStartElement( "Department" )
                $writer.WriteElementString( "Name", $department.Name )
                $writer.WriteElementString( "Facility", $department.Facility )
                foreach($position in $department.Positions){
                    $writer.WriteStartElement( "Position" )
                    $writer.WriteElementString( "Name", $position.Name )
                    foreach($name in $position.Names){
                        $writer.WriteElementString( "Nev", $name )
                    }
                    $writer.WriteEndElement()
                }
                $writer.WriteEndElement()
            }
            $writer.WriteEndElement()
            $writer.WriteEndDocument()
        }
        finally {
            if($writer){
                $writer.Close()
            }
        }
    }
}