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

function Set-DuplexingMode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Simplex", "Duplex")]
        [string]$Mode
    )

    process {
        Write-Log -Message "Setting duplexing mode to: $Mode..."

        # Relatív útvonal meghatározása a Private mappához
        $configDir = Join-Path $PSScriptRoot "..\Private"
        $configFile = Join-Path $configDir "$($Mode.ToLower()).bin"
        
        if (-not (Test-Path $configFile)) {
            throw "Configuration file not found: $configFile"
        }

        try {
            # 1. Alapértelmezett nyomtató azonosítása
            $printerSettings = New-Object System.Drawing.Printing.PrinterSettings
            $printerName = $printerSettings.PrinterName
            $regPath = "HKCU:\Printers\DevModePerUser"
            
            $printerKey = (Get-ItemProperty $regPath | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -like "*$printerName*" }).Name

            if (-not $printerKey) {
                throw "Printer ($printerName) not found in Registry!"
            }

            # 2. Bináris adatok beírása a Registry-be
            $binData = Get-Content $configFile -AsByteStream
            Set-ItemProperty -Path $regPath -Name $printerKey -Value $binData
            
            # 3. Értesítés küldése a Windowsnak (SendMessageTimeout)
            $signature = @"
            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
"@
            if (-not ([System.Management.Automation.PSTypeName]"Win32.Win32SendMessage").Type) {
                Add-Type -MemberDefinition $signature -Name "Win32SendMessage" -Namespace Win32 -ErrorAction SilentlyContinue
            }

            $WM_SETTINGCHANGE = 0x001A
            $result = [IntPtr]::Zero
            [Win32.Win32SendMessage]::SendMessageTimeout([IntPtr]0xffff, $WM_SETTINGCHANGE, [IntPtr]::Zero, "Printers", 0x0002, 1000, [ref]$result)

            Write-Log -Message "Mode $Mode is set for printer: $printerName"
        }
        catch {
            # Hibát csak továbbdobjuk, a hívó kezeli és logolja
            throw
        }
    }
}

