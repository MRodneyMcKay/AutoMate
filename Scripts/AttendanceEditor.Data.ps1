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
