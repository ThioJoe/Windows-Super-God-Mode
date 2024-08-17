# This script will find text strings of "ms-settings:" in a DLL file and output them to a text file.
# Meant to be run on "SystemSettings.dll" in C:\Windows\ImmersiveControlPanel\
#
# Optional Arguments:
#    -DllPath: Path to the DLL file to search
#    -OutputFilePath: Path to the output text file
#
# Without arguments, the script will prompt the user for the DLL file path, and will output text file result to same directory as script
#
# Example Usage:
#    .\Get-MS-Settings-Strings.ps1 -DllPath "C:\Windows\ImmersiveControlPanel\SystemSettings.dll" -OutputFilePath "SystemSettings-MS-Settings.txt"
#

param (
    [string]$DllPath,
    [string]$OutputFilePath
)

function Get-DllMsSettings {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DllPath
    )

    if (-not (Test-Path $DllPath)) {
        Write-Error "File not found: $DllPath"
        return @()
    }

    $content = [System.IO.File]::ReadAllText($DllPath, [System.Text.Encoding]::Unicode)
    $results = New-Object System.Collections.Generic.HashSet[string]

    $matches = [regex]::Matches($content, 'ms-settings:[a-z-]+')
    foreach ($match in $matches) {
        [void]$results.Add($match.Value)
    }

    Write-Host "Unique Matches Found: $($results.Count)"
    return $results | Sort-Object
}

# If no parameter for DLL path is provided, prompt the user
if (-not $DllPath) {
    Write-Host "`nEnter the path to the DLL file. Or press enter to use default path: C:\Windows\ImmersiveControlPanel\SystemSettings.dll"
    $DllPath = Read-Host "`nEnter Path"
    if (-not $DllPath) {
        Write-Host "Using default path: C:\Windows\ImmersiveControlPanel\SystemSettings.dll"
        $DllPath = "C:\Windows\ImmersiveControlPanel\SystemSettings.dll"
    }
}
# Check the path of the DLL file
if (-not (Test-Path $DllPath)) {
    Write-Error "File not found: $DllPath"
    return
}

# If no output path argument is given, set text file based on input file name, in same directory as script working directory
if (-not $OutputFilePath) {
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($DllPath)
    $OutputFilePath = [System.IO.Path]::Combine($PSScriptRoot, "$fileName-MS-Settings.txt")
} else {
    # Check if it's a relative or absolute path, and if relative then make it relative to the script
    if (-not [System.IO.Path]::IsPathRooted($OutputFilePath)) {
        $OutputFilePath = [System.IO.Path]::Combine($PSScriptRoot, $OutputFilePath)
    }
}

# Call main function
Write-Host "`nBeginning search...`n"
$results = Get-DllMsSettings -DllPath $DllPath

# Output the results to a txt file
$results | Out-File -FilePath $OutputFilePath

Write-Host "Results written to file: $OutputFilePath`n"