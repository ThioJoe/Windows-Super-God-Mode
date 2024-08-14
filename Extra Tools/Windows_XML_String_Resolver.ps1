# Windows XML String Resolver
# Author: ThioJoe
#
# Purpose: This script takes an XML file that contains string references (e.g., "@shell32.dll,-1234")
#          and resolves them to their actual string values. It's particularly useful for working with
#          Windows resource files or any XML that uses similar string reference formats.
#
# How to Use:
# 1. Open PowerShell and navigate to the path containing this script using the 'cd' command.
# 2. Run the following command to allow running scripts for the current session:
#        Set-ExecutionPolicy -ExecutionPolicy unrestricted -Scope Process
# 3. Without closing the PowerShell window, run the script by typing the name of the script file starting with .\ for example:
#        .\Windows_XML_String_Resolver.ps1

# ------------------------- ARGUMENTS -------------------------
# -XmlFilePath
#     String (Required)
#     The path to the XML file that needs to be processed. Can be a relative or absolute path.
#
# -CustomResourcePaths
#     String Array (Optional)
#     Specify custom paths for DLL or MUI files. Each entry should be in the format "dllName=path".
#     The dllName can be with or without the .dll extension and is case-insensitive.
#     Example: "shell32=C:\custom\path\shell32.dll", "user32=C:\another\path\user32.mui"
#
# -Debug
#     Switch (Takes no values)
#     Enable debug output for more detailed information during script execution.
#     Shows the full paths of DLLs being loaded and lists all custom resource paths.
#
# ---------------------------------------------------------------------
#
#   EXAMPLE USAGE FROM COMMAND LINE:
#       .\Windows_XML_String_Resolver.ps1 -XmlFilePath "path\to\your\file.xml" -CustomResourcePaths "shell32=C:\custom\path\shell32.dll", "user32=C:\another\path\user32.mui" -Debug
#
# ---------------------------------------------------------------------

param(
    [string]$XmlFilePath,
    [string[]]$CustomResourcePaths,
    [switch]$Debug
)

# Import necessary .NET classes
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class Windows {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool FreeLibrary(IntPtr hModule);
}
"@

# Initialize a case-insensitive hashtable to store custom DLL paths
$customDllPathMap = New-Object System.Collections.Hashtable([System.StringComparer]::OrdinalIgnoreCase)

# Parse custom DLL paths
if ($CustomResourcePaths) {
    foreach ($pair in $CustomResourcePaths) {
        $parts = $pair -split '='
        if ($parts.Length -eq 2) {
            $dllName = $parts[0].Trim().TrimEnd('.dll')  # Remove .dll if present
            $dllPath = $parts[1].Trim()
            $customDllPathMap[$dllName] = $dllPath
        }
        else {
            Write-Warning "Invalid custom DLL path format: $pair. Expected format: 'dllName=path'"
        }
    }
}

function Get-LocalizedString {
    param ( [string]$StringReference )

    if ($StringReference -match '@(.+),-(\d+)') {
        $dllName = $Matches[1].TrimEnd('.dll')  # Remove .dll if present
        $resourceId = [uint32]$Matches[2]

        # Check if we have a custom path for this DLL
        if ($customDllPathMap.ContainsKey($dllName)) {
            $dllPath = $customDllPathMap[$dllName]
        }
        else {
            $dllPath = [Environment]::ExpandEnvironmentVariables($Matches[1])
        }

        if ($Debug) {
            Write-Host "Loading DLL: $dllPath" -ForegroundColor Cyan
        }

        $hModule = [Windows]::LoadLibraryEx($dllPath, [IntPtr]::Zero, 0x00000800) # LOAD_LIBRARY_AS_DATAFILE
        if ($hModule -eq [IntPtr]::Zero) {
            $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Warning "Failed to load library: $dllPath. Error code: $errorCode"
            return $StringReference
        }

        $stringBuilder = New-Object System.Text.StringBuilder 1024
        $result = [Windows]::LoadString($hModule, $resourceId, $stringBuilder, $stringBuilder.Capacity)

        [void][Windows]::FreeLibrary($hModule)

        if ($result -ne 0) {
            return $stringBuilder.ToString()
        } else {
            $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Warning "Failed to load string resource: $resourceId from $dllPath. Error code: $errorCode"
            return $StringReference
        }
    } else {
        return $StringReference
    }
}

function Resolve-XmlStringReferences {
    param (
        [string]$XmlContent
    )

    $resolvedXml = $XmlContent

    $stringRefPattern = '@[^,]+,-\d+'
    $stringMatches = [regex]::Matches($XmlContent, $stringRefPattern)

    foreach ($match in $stringMatches) {
        $originalString = $match.Value
        $resolvedString = Get-LocalizedString $originalString
        if ($resolvedString -ne $originalString) {
            $resolvedXml = $resolvedXml.Replace($originalString, [System.Security.SecurityElement]::Escape($resolvedString))
        }
    }

    return $resolvedXml
}

# Main script logic
if (-not $XmlFilePath) {
    $XmlFilePath = Read-Host "Please enter the path to the XML file"
}

# Remove quotes if present
$XmlFilePath = $XmlFilePath.Trim('"')

if (-not (Test-Path $XmlFilePath)) {
    Write-Error "The specified file does not exist: $XmlFilePath"
    exit 1
}

if ($Debug) {
    Write-Host "Custom DLL Paths:" -ForegroundColor Yellow
    $customDllPathMap.GetEnumerator() | ForEach-Object {
        Write-Host "$($_.Key) = $($_.Value)" -ForegroundColor Yellow
    }
}

try {
    $xmlContent = Get-Content $XmlFilePath -Raw
    $resolvedXml = Resolve-XmlStringReferences $xmlContent

    $outputPath = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($XmlFilePath),
        [System.IO.Path]::GetFileNameWithoutExtension($XmlFilePath) + "-resolved.xml"
    )

    $resolvedXml | Out-File $outputPath -Encoding UTF8
    Write-Host "Resolved XML saved to: $outputPath"
} catch {
    Write-Error "An error occurred: $_"
    exit 1
}