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

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xml.Linq

$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules"
$ResolvedModulePath = (Resolve-Path -Path $ModulePath).Path
Import-Module (Join-Path -Path $ResolvedModulePath -ChildPath "LoggingSystem\LoggingSystem.psd1")

. (Join-Path -Path $PSScriptRoot -ChildPath "AttendanceEditor.Data.ps1")
. (Join-Path -Path $PSScriptRoot -ChildPath "AttendanceEditor.View.ps1")
. (Join-Path -Path $PSScriptRoot -ChildPath "AttendanceEditor.Logic.ps1")

Write-Log -Message "Jelenléti szerkesztő indul (XML)" -Level "INFO"

Initialize-EditorState
$Window = Initialize-EditorWindow

Load-All

try {
    $Window.ShowDialog() | Out-Null
}
catch [System.InvalidOperationException] {
    if ($_.Exception.Message -match 'after a Window has closed') {
        Initialize-EditorState
        $Window = Initialize-EditorWindow
        Load-All
        $Window.ShowDialog() | Out-Null
    }
    else {
        throw
    }
}
