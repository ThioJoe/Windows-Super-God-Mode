# Checks if current PowerShell environment already has 'Win32' type defined, and if not add definition for it
# Otherwise it will throw an error if it was already added like if the script is re-ran without closing the Window sometimes
# The 'Win32' type provides access to some key Windows API functions
if (-not ([System.Management.Automation.PSTypeName]'Win32').Type) {
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class Win32 {
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
    public static extern int SHLoadIndirectString(string pszSource, StringBuilder pszOutBuf, int cchOutBuf, IntPtr ppvReserved);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr LoadLibrary(string lpFileName);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);
}
"@
}

function Get-LocalizedString {
    param (
        [string]$StringReference
    )
    if ($StringReference -match '@\{.+\?ms-resource://.+}') {
        return Get-MsResource $StringReference
    }
    elseif ($StringReference -match '@(.+),-(\d+)') {
        $dllPath = [Environment]::ExpandEnvironmentVariables($Matches[1])
        $resourceId = [uint32]$Matches[2]
        return Get-StringFromDll $dllPath $resourceId
    }
    else {
        Write-Error "Invalid string reference format: $StringReference"
        return
    }
}

function Get-MsResource {
    param (
        [string]$ResourcePath
    )
    $stringBuilder = New-Object System.Text.StringBuilder 1024
    $result = [Win32]::SHLoadIndirectString($ResourcePath, $stringBuilder, $stringBuilder.Capacity, [IntPtr]::Zero)
    if ($result -eq 0) {
        return $stringBuilder.ToString()
    } else {
        Write-Error "Failed to retrieve ms-resource: $ResourcePath. Error code: $result"
        return $null
    }
}

function Get-StringFromDll {
    param (
        [string]$DllPath,
        [uint32]$ResourceId
    )
    # Calls the 'LoadLibrary' method from the 'Win32' class defined earlier. Loads the DLL containing the reference
    $hModule = [Win32]::LoadLibrary($DllPath)
    if ($hModule -eq [IntPtr]::Zero) {
        Write-Error "Failed to load library: $DllPath"
        return
    }

    $stringBuilder = New-Object System.Text.StringBuilder 1024
    $result = [Win32]::LoadString($hModule, $ResourceId, $stringBuilder, $stringBuilder.Capacity)

    if ($result -ne 0) {
        return $stringBuilder.ToString()
    } else {
        Write-Error "Failed to load string resource: $ResourceId from $DllPath"
    }
}

Write-Host "Enter 'x' at any time to exit the program."
while ($true) {
    Write-Host "`n------------------------------------------------------------------------"
    Write-Host "Enter the string resource reference to get."
    Write-Host " > Example 1: @%SystemRoot%\system32\shell32.dll,-9227"
    Write-Host " > Example 2: @{windows?ms-resource://Windows.UI.SettingsAppThreshold/SearchResources/SystemSettings_CapabilityAccess_Gaze_UserGlobal/Description}"
    Write-Host "`nResource Reference:  " -NoNewline
    $userInput = Read-Host
    if ($userInput.ToLower() -eq 'x') {
        Write-Host "Exiting the program. Goodbye!"
        break
    }
    $localizedString = Get-LocalizedString $userInput
    if ($localizedString) {
        Write-Host "`n   Returned Value: " -NoNewline
        Write-Host $localizedString -ForegroundColor Yellow
    }
}