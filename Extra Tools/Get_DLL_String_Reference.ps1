# Checks if current PowerShell environment already has 'Win32' type defined, and if not add definition for it
# Otherwise it will throw an error if it was already added like if the script is re-ran without closing the Window sometimes
# The 'Win32' type provides access to some key Windows API functions
if (-not ([System.Management.Automation.PSTypeName]'Win32').Type) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    public class Win32 {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr LoadLibrary(string lpFileName);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int LoadString(IntPtr hInstance, int uID, StringBuilder lpBuffer, int nBufferMax);
    }
"@
}

function Get-LocalizedString {
    param (
        [string]$StringReference
    )
    
	# Separates out the various parts of the string reference based on regex pattern
    if ($StringReference -match '@(.+),-(\d+)') {
        $dllPath = [Environment]::ExpandEnvironmentVariables($Matches[1])
        $resourceId = [int]$Matches[2]
		# Calls the 'LoadLibrary' method from the 'Win32' class defined earlier. Loads the DLL containing the reference
        $hModule = [Win32]::LoadLibrary($dllPath)
        if ($hModule -eq [IntPtr]::Zero) {
            Write-Error "Failed to load library: $dllPath"
            return
        }
        
        $stringBuilder = New-Object System.Text.StringBuilder 1024
		# Calls the 'LoadString' method from the 'Win32' class, to retrieve string resource with specified ID from the loaded DLL
        $result = [Win32]::LoadString($hModule, $resourceId, $stringBuilder, $stringBuilder.Capacity)
        
        if ($result -ne 0) {
            return $stringBuilder.ToString()
        } else {
            Write-Error "Failed to load string resource: $resourceId from $dllPath"
        }
    } else {
        Write-Error "Invalid string reference format: $StringReference"
    }
}

Write-Host "Enter 'x' at any time to exit the program."

while ($true) {
    # Prompt the user for input
    Write-Host "`n------------------------------------------------------------------------"
    Write-Host "Enter the full path and index of the string resource to get."
    Write-Host " > Example: @%SystemRoot%\system32\shell32.dll,-9227"
    Write-Host "`nPath and Index:  " -NoNewline
    $userInput = Read-Host

    if ($userInput.ToLower() -eq 'x') {
        Write-Host "Exiting the program. Goodbye!"
        break
    }

    # Get and display the localized string
    $localizedString = Get-LocalizedString $userInput
    if ($localizedString) {
        Write-Host "`n   Returned Value: " -NoNewline
        Write-Host $localizedString -ForegroundColor Yellow
    }
}