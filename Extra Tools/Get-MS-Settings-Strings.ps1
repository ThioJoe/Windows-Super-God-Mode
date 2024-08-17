# This script will find text strings of "ms-settings:" in a DLL file and output them to a text file.
# Meant to be run on "SystemSettings.dll" in C:\Windows\ImmersiveControlPanel\
#
# Optional Arugments:
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

    # Check if the file exists
    if (-not (Test-Path $DllPath)) {
        Write-Error "File not found: $DllPath"
        return @()
    }

    # Read the DLL file as bytes
    $bytes = [System.IO.File]::ReadAllBytes($DllPath)

    $searchString = "ms-settings:"
    $searchBytes = [System.Text.Encoding]::Unicode.GetBytes($searchString)
    $results = @()
    $matchCount = 0

    for ($i = 0; $i -lt $bytes.Length - $searchBytes.Length; $i += 2) {
        $potentialMatch = [System.Text.Encoding]::Unicode.GetString($bytes, $i, $searchBytes.Length)
        if ($potentialMatch -like "*ms-settings:*") {
            $end = $i
            while ($end -lt $bytes.Length - 1 -and -not ($bytes[$end] -eq 0 -and $bytes[$end + 1] -eq 0)) {
                $end += 2
            }
            $length = $end - $i
            $result = [System.Text.Encoding]::Unicode.GetString($bytes, $i, $length)
            $results += $result
            $matchCount++
            Write-Host "Match found: $result"
            $i = $end
        }
    }

    Write-Host "Total matches found: $matchCount"
    $uniqueResults = $results | Select-Object -Unique
    Write-Host "Unique matches: $($uniqueResults.Count)"

    # Return unique results
    return $uniqueResults
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
Write-Host "`nBeginning search..."
$results = Get-DllMsSettings -DllPath $DllPath
Write-Host "`nResults found: $($results.Count)"

# Sort the results
$results = $results | Sort-Object

# Output the results to a txt file
$results | Out-File -FilePath $OutputFilePath

Write-Host "Results written to file: $OutputFilePath"
