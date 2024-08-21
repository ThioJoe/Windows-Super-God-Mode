# This script fetches the URI protocols for each installed AppxPackage via their AppxManifest.xml file, then brute force searches for those URIs in all files in the app's install directory.

# Function to get app details including URI protocols and install paths
function Get-AppDetails {
    $result = [System.Collections.ArrayList]@()
    foreach ($appx in Get-AppxPackage) {
        $location = $appx.InstallLocation
        $manifest = "$location\AppxManifest.xml"
        if ($null -ne $location -and (Test-Path $manifest -PathType Leaf)) {
            [xml]$xml = Get-Content $manifest
            $ns = New-Object Xml.XmlNamespaceManager $xml.NameTable
            $ns.AddNamespace("main", "http://schemas.microsoft.com/appx/manifest/foundation/windows10")
            $ns.AddNamespace("uap", "http://schemas.microsoft.com/appx/manifest/uap/windows10")
            $ns.AddNamespace("uap2", "http://schemas.microsoft.com/appx/manifest/uap/windows10/2")
            $ns.AddNamespace("uap3", "http://schemas.microsoft.com/appx/manifest/uap/windows10/3")
            $ns.AddNamespace("uap4", "http://schemas.microsoft.com/appx/manifest/uap/windows10/4")
            $ns.AddNamespace("uap5", "http://schemas.microsoft.com/appx/manifest/uap/windows10/5")

            $uapNamespaces = @("uap", "uap2", "uap3", "uap4", "uap5")
            $uriXpathQuery = ($uapNamespaces | ForEach-Object {
                "//$_`:Extension[@Category = 'windows.protocol']/$_`:Protocol/@Name"
            }) -join ' | '

            $uris = $xml.SelectNodes($uriXpathQuery, $ns) | Select-Object -ExpandProperty '#text'

            if ($uris.Count -gt 0) {
                $tmp = [PSCustomObject]@{
                    Name = $appx.Name
                    URIs = $uris
                    Folder = $appx.InstallLocation
                }
                $null = $result.Add($tmp)
            }
        }
    }
    return $result
}

# Define encoding mappings for different file extensions
$encodingMap = @{
    ".txt"  = "UTF-8"
    ".xml"  = "UTF-8"
    ".json" = "UTF-8"
    ".dll"  = "Unicode"
    ".exe"  = "Unicode"
    ".js"   = "UTF-8"
    ".map"  = "UTF-8"
    # Add more mappings as needed
}

function Get-ProtocolsInFile {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$protocolsList,
        [Parameter(Mandatory=$true)]
        [string]$filePathToCheck,
        [Parameter(Mandatory=$true)]
        [hashtable]$encodingMap
    )

    if (-not (Test-Path $filePathToCheck)) {
        Write-Error "File not found: $filePathToCheck"
        return $null
    }

    $results = @{}
    $fileExtension = [System.IO.Path]::GetExtension($filePathToCheck).ToLower()
    $encodingsToTry = @()

    if ($encodingMap.ContainsKey($fileExtension)) {
        $encodingsToTry += $encodingMap[$fileExtension]
    } else {
        $encodingsToTry += "UTF-8", "Unicode"
    }

    foreach ($encodingName in $encodingsToTry) {
        $encoding = [System.Text.Encoding]::GetEncoding($encodingName)

        try {
            $content = [System.IO.File]::ReadAllText($filePathToCheck, $encoding)

            foreach ($protocol in $protocolsList) {
                # Different patterns for UTF-8 and Unicode
                if ($encodingName -eq "UTF-8") {
                    $uriPattern = [regex]::Escape($protocol) + "://[^""\s<>()\\``]+"
                } else {
                    $uriPattern = [regex]::Escape($protocol) + "://[\x20-\x7E]+"
                }

                $matches = [regex]::Matches($content, $uriPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

                if ($matches.Count -gt 0) {
                    # Store the matches along with the encoding used and other data
                    $results[$protocol] = @($matches | ForEach-Object {
                        # If the match contains a bracket, or ends with an equals sign, mark it as "UsesVariables"
                        $usesVariables = $_.Value -match "[<>()\[\]]|=$"

                        [PSCustomObject]@{
                            FullURL = $_.Value
                            EncodingUsed = $encodingName
                            UsesVariables = $usesVariables
                        }
                    })
                }
            }

            # If we found matches, no need to try other encodings
            if ($results.Count -gt 0) {
                break
            }
        }
        catch {
            Write-Warning "Error processing file $filePathToCheck with $encodingName encoding: $_"
        }
    }

    return @{
        FilePath = $filePathToCheck
        Matches = $results
    }
}


function OutputCSV {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.ArrayList]$data,
        [Parameter(Mandatory=$true)]
        [string]$outputPath
    )

    try {
        $data | Export-Csv -Path $outputPath -NoTypeInformation
        Write-Host "Results exported to: $outputPath"
    }
    catch {
        Write-Error "Error exporting to CSV: $_"
        $outputPath = Read-Host "`nEnter the path to save the CSV file"
        $outputPath = $outputPath.Trim('"')
        if (-not (Test-Path $outputPath)) {
            New-Item -Path $outputPath -ItemType File -Force | Out-Null
        }
        $data | Export-Csv -Path $outputPath -NoTypeInformation
        Write-Host "If there was no error, results exported to: $outputPath"
    }
}

# Main script execution
$appDetails = Get-AppDetails

$results = @()
$searchedFiles = @()

# Define ignored file extensions
$ignoredExtensions = @(
    # Images
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', ".ico",
    # Other irrelevant file types
    ".p7x", ".ttf", ".onnxe",
    # Compressed or other files that don't produce reliable results
    ".bundle", '.vsix'
)

# Main script execution
$appDetails = Get-AppDetails

$results = @()
$searchedFiles = @()

$totalFiles = ($appDetails | ForEach-Object {
    Get-ChildItem -Path $_.Folder -Recurse -File | Where-Object { $_.Extension -notin $ignoredExtensions }
}).Count

$processedFiles = 0
$lastPercentage = -1  # Initialize to -1 to ensure the first 0% is displayed
$processedFiles = 0
$currentPercentage = 0

foreach ($app in $appDetails) {
    Write-Verbose "`rSearching in $($app.Name) | For URIs: $($app.URIs -join ', ')"
    $files = Get-ChildItem -Path $app.Folder -Recurse -File | Where-Object { $_.Extension -notin $ignoredExtensions }
    foreach ($file in $files) {
        $searchedFiles += $file.FullName
        $fileResults = Get-ProtocolsInFile -protocolsList $app.URIs -filePathToCheck $file.FullName -encodingMap $encodingMap
        if ($fileResults -and $fileResults.Matches.Count -gt 0) {
            $results += $fileResults
        }
        $processedFiles++
        $currentPercentage = [math]::Floor(($processedFiles / $totalFiles) * 100)
        if ($currentPercentage -ne $lastPercentage) {
            Write-Host "`rProgress: $currentPercentage%" -NoNewline
            $lastPercentage = $currentPercentage
        }
    }
}

Write-Host "`nProcessing complete.$(" " * $paddingLength)"

# Prepare data for CSV export
$csvData = @()
foreach ($result in $results) {
    if ($result -and $result.Matches) {
        foreach ($protocol in $result.Matches.Keys) {
            foreach ($match in $result.Matches[$protocol]) {
                $csvData += [PSCustomObject]@{
                    FilePath = $result.FilePath
                    Protocol = $protocol
                    FullURL = $match.FullURL
                    EncodingUsed = $match.EncodingUsed
                    UsesVariables = $match.UsesVariables
                }
            }
        }
    }
}

# Create the output directory if it doesn't exist
$outputDir = ".\ProtocolMatches"
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

# Get date to use in the output file names
$date = Get-Date -Format 'yyyyMMdd_HHmmss'

# Export searched files list
$searchedFilesPath = Join-Path $outputDir "files_searched_$date.txt"
$searchedFiles | Out-File -FilePath $searchedFilesPath -Encoding utf8

# Export to CSV
$outputPath = Join-Path $outputDir "protocol_matches_$date.csv"

OutputCSV -data $csvData -outputPath $outputPath

Write-Host "Searched files list exported to: $searchedFilesPath"
Write-Host "Protocol matches exported to: $outputPath"
