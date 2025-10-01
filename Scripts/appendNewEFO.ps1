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

# Import your logging system
$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\Modules\"
$resolvedModulegPath = (Resolve-Path -Path $ModulePath).Path
Import-Module (Join-Path -Path $resolvedModulegPath -ChildPath 'LoggingSystem\LoggingSystem.psd1')

# Add WPF GUI for input
Add-Type -AssemblyName PresentationFramework

# Define WPF window as XAML
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Add to EfoNévsor" Height="150" Width="400" WindowStartupLocation="CenterScreen">
    <StackPanel Margin="10">
        <TextBlock Text="Add a new name to EfoNévsor:" Margin="0,0,0,10" FontWeight="Bold"/>
        <TextBox Name="InputBox" Height="25" Margin="0,0,0,10"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="OKButton" Width="75" Margin="5">OK</Button>
            <Button Name="CancelButton" Width="75" Margin="5">Cancel</Button>
        </StackPanel>
    </StackPanel>
</Window>
"@

# Load XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Reference UI elements
$inputBox   = $window.FindName("InputBox")
$okButton   = $window.FindName("OKButton")
$cancelBtn  = $window.FindName("CancelButton")

# Wire up buttons
$okButton.Add_Click({
    $window.DialogResult = $true
    $window.Close()
})
$cancelBtn.Add_Click({
    $window.DialogResult = $false
    $window.Close()
})

# Show window and handle result
if ($window.ShowDialog() -eq $true) {
    $newName = $inputBox.Text.Trim()
    if ($newName) {
        $existing = [System.Environment]::GetEnvironmentVariable("EfoNévsor", "User")
        if ([string]::IsNullOrEmpty($existing)) {
            $updated = $newName
        } else {
            $updated = "$existing;$newName"
        }

        [System.Environment]::SetEnvironmentVariable("EfoNévsor", $updated, "User")
        Write-Log -Message "'$newName'fel került az EFO névsorba" -Level INFO
        [System.Windows.MessageBox]::Show("'$newName'fel került az EFO névsorba", "Siker")
    } else {
        Write-Log -Message "Nem került új név hozzáadásra" -Level INFO
        [System.Windows.MessageBox]::Show("Nem került új név hozzáadásra", "Info")
    }
} else {
    Write-Log -Message "Nem került új név hozzáadásra" -Level INFO
}
