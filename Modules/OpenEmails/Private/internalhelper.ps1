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

function Open-EmailTemplate {
    param (
        [object]$Outlook,
        [string]$TemplatePath,
        [hashtable]$Replacements,
        [string]$subject
    )
    $msg = $Outlook.CreateItemFromTemplate($TemplatePath)
    foreach ($key in $Replacements.Keys) {
        $msg.HTMLBody = $msg.HTMLBody.Replace($key, $Replacements[$key])
    }
    $inspector = $msg.GetInspector()
    if ($inspector) {
        Write-Log -Message "$subject email opened."
    } else {
        Write-Log -Message "$subject could not be opened." -Level "ERROR"
    }
    $inspector.Display()
}

$code = @"
using System;
using System.Runtime.InteropServices;

public static class InteropCom
{
    // Use CLSIDFromProgIDEx to convert "Outlook.Application" into its CLSID.
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgIDEx(string progId, out Guid clsid);
    
    // Call GetActiveObject from oleaut32.dll to get the running instance.
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    public static object GetActiveInstance(string progId, bool throwOnError)
    {
        if (string.IsNullOrWhiteSpace(progId))
            throw new ArgumentNullException(nameof(progId));

        int hr = CLSIDFromProgIDEx(progId, out Guid clsid);
        if (hr < 0)
        {
            if (throwOnError)
                Marshal.ThrowExceptionForHR(hr);
            return null;
        }

        hr = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
        if (hr < 0)
        {
            if (throwOnError)
                Marshal.ThrowExceptionForHR(hr);
            return null;
        }
        return obj;
    }
}
"@

Add-Type -TypeDefinition $code -Language CSharp -PassThru