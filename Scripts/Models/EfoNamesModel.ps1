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

class EfoNamesModel {
    [ObservableCollection[string]] $Names
    [string] $XmlPath

    EfoNamesModel([string] $xmlPath) {
        $this.XmlPath = $xmlPath
        $this.Names   = [ObservableCollection[string]]::new()
        foreach ($name in $this.Read()) {
            $this.Names.Add($name)
        }
    }

    [string[]] Read() {
        if (-not (Test-Path -LiteralPath $this.XmlPath)) {
            return @()
        }
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo("hu-HU")
        try {
            $xml = [xml](Get-Content -LiteralPath $this.XmlPath)

            return @(
                $xml.EFONames.Name |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Culture $culture
            )
        }
        catch {
            Write-Log `
                -Message "EFO XML beolvasási hiba: $($_.Exception.Message)" `
                -Level "ERROR"
            return @()
        }
    }

    [void] Sort() {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo("hu-HU")

        $sorted = @( $this.Names | Sort-Object -Culture $culture )
        $this.Names.Clear()
        foreach ($name in $sorted) {
            $this.Names.Add($name)
        }
    }

    [bool] Save() {
        try {
            $culture = [System.Globalization.CultureInfo]::GetCultureInfo("hu-HU")
            $xml  = New-Object System.Xml.XmlDocument
            $root = $xml.CreateElement("EFONames")
            $xml.AppendChild($root) | Out-Null
            $sorted = @( $this.Names | Sort-Object -Culture $culture )
            foreach ($name in $sorted) {
                $node = $xml.CreateElement("Name")
                $node.InnerText = $name
                $root.AppendChild($node) | Out-Null
            }

            $xml.Save($this.XmlPath)
            return $true
        }
        catch {
            Write-Log -Message "EFO XML mentési hiba: $($_.Exception.Message)" -Level "ERROR"
            return $false
        }
    }

    [void] Reload() {
        $this.Names.Clear()

        foreach ($name in $this.Read()) {
            $this.Names.Add($name)
        }
    }
}
