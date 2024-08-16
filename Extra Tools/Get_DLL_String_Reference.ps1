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
    [CmdletBinding()]
    param (
        [string]$StringReference
    )
    if ($StringReference -match '@\{.+\?ms-resource://.+}') {
        return Get-MsResource $StringReference -Verbose:$VerbosePreference
    }
    elseif ($StringReference -match '@(.+),-(\d+)') {
        $dllPath = [Environment]::ExpandEnvironmentVariables($Matches[1])
        $resourceId = [uint32]$Matches[2]
        return Get-StringFromDll $dllPath $resourceId -Verbose:$VerbosePreference
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
    Write-Verbose "Attempting to retrieve resource: $ResourcePath"
    $stringBuilder = New-Object System.Text.StringBuilder 1024
    $result = [Win32]::SHLoadIndirectString($ResourcePath, $stringBuilder, $stringBuilder.Capacity, [IntPtr]::Zero)
    if ($result -eq 0) {
        Write-Verbose "Successfully retrieved resource on first attempt"
        return $stringBuilder.ToString()
    } else {
        Write-Verbose "Initial attempt failed with error code: $result. Trying alternative methods..."

        # Extract package name and resource URI
        $packageFullName = ($ResourcePath -split '\?')[0].Trim('@{}')
        $resourceUri = ($ResourcePath -split '\?')[1]
        Write-Verbose "Extracted package full name: $packageFullName"
        Write-Verbose "Extracted resource URI: $resourceUri"

        # Extract package name without version and architecture
        $packageName = ($packageFullName -split '_')[0]
        Write-Verbose "Extracted package name: $packageName"

        # Find the package installation path
        $package = Get-AppxPackage | Where-Object { $_.Name -eq $packageName }
        if (-not $package) {
            # If exact match fails, try matching by package family name
            $packageFamilyName = ($packageFullName -split '_')[-1]
            $package = Get-AppxPackage | Where-Object { $_.PackageFamilyName -eq "${packageName}_$packageFamilyName" }
        }

        if ($package) {
            $packagePath = $package.InstallLocation
            Write-Verbose "Package installation path: $packagePath"
            $priPath = Join-Path $packagePath "resources.pri"
            Write-Verbose "Attempting to use resources.pri at: $priPath"
            if (Test-Path $priPath) {
                $newResourcePath = "@{" + $priPath + "?" + $resourceUri
                Write-Verbose "New resource path: $newResourcePath"
                $result = [Win32]::SHLoadIndirectString($newResourcePath, $stringBuilder, $stringBuilder.Capacity, [IntPtr]::Zero)
                if ($result -eq 0) {
                    Write-Verbose "Successfully retrieved resource using resources.pri"
                    return $stringBuilder.ToString()
                }
                Write-Error "Failed to retrieve using resources.pri. Error code: $result"
            } else {
                Write-Verbose "resources.pri not found at expected location"
            }
        } else {
            Write-Verbose "Package not found"
        }

        # If still failed, try without the /resources/ folder
        $resourceUriWithoutResources = $resourceUri -replace '/resources/', '/'
        $newResourcePath = "@{" + $priPath + "?" + $resourceUriWithoutResources
        Write-Verbose "Attempting without /resources/ folder. New path: $newResourcePath"
        $result = [Win32]::SHLoadIndirectString($newResourcePath, $stringBuilder, $stringBuilder.Capacity, [IntPtr]::Zero)
        if ($result -eq 0) {
            Write-Verbose "Successfully retrieved resource without /resources/ folder"
            return $stringBuilder.ToString()
        }
        Write-Host "Failed to retrieve without /resources/ folder. Error code: $result"

        Write-Error "Failed to retrieve ms-resource: $ResourcePath. Error code: $result"
        return $null
    }
}

function Get-StringFromDll {
    [CmdletBinding()]
    param (
        [string]$DllPath,
        [uint32]$ResourceId
    )
    Write-Verbose "Attempting to load string from DLL: $DllPath, Resource ID: $ResourceId"
    $hModule = [Win32]::LoadLibrary($DllPath)
    if ($hModule -eq [IntPtr]::Zero) {
        Write-Error "Failed to load library: $DllPath"
        return
    }

    $stringBuilder = New-Object System.Text.StringBuilder 1024
    $result = [Win32]::LoadString($hModule, $ResourceId, $stringBuilder, $stringBuilder.Capacity)

    if ($result -ne 0) {
        Write-Verbose "Successfully loaded string from DLL"
        return $stringBuilder.ToString()
    } else {
        Write-Error "Failed to load string resource: $ResourceId from $DllPath"
    }
}

Write-Host "Enter 'x' at any time to exit the program."
while ($true) {
    Write-Verbose "Verbose Mode."
    Write-Host "`n------------------------------------------------------------------------"
    Write-Host "Enter the string resource reference to get."
    Write-Host " > Example 1: @%SystemRoot%\system32\shell32.dll,-9227"
    Write-Host " > Example 2: @{windows?ms-resource://Windows.UI.SettingsAppThreshold/SearchResources/SystemSettings_CapabilityAccess_Gaze_UserGlobal/Description}"
    Write-Host " > Example 3: @{Microsoft.SecHealthUI_8wekyb3d8bbwe?ms-resource://Microsoft.SecHealthUI/Resources/AccountTileMenuEntryAndTitle}"
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