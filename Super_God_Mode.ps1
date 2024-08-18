# Get All Shell Folder Shortcuts Script (Aka "Super God Mode")
#
# This PowerShell script is designed to find and create shortcuts for all special shell folders in Windows.
# These folders can be identified through their unique Class Identifiers (CLSIDs) or by their names.
# The script also generates CSV files listing these folders and associated tasks/links.

# How to Use:
# 1. Open PowerShell and navigate to the path containing this script using the 'cd' command.
# 2. Run the following command to allow running scripts for the current session:
#        Set-ExecutionPolicy -ExecutionPolicy unrestricted -Scope Process
# 3. Without closing the PowerShell window, run the script by typing the name of the script file starting with  .\  for example:
#        .\Super_God_Mode.ps1
# 4. Wait for it to finish, then check the "Shell Folder Shortcuts" folder for the generated shortcuts.

# ------------------------- OPTIONAL ARGUMENTS -------------------------
# -SaveCSV
#     Switch (Takes no values)
#     Save info about all the shortcuts into CSV spreadsheet files for each type of shell folder
#
# -SaveXML
#     Switch (Takes no values)
#     Save the XML content from shell32.dll as a file containing info about the task links
#
# -Output
#     String (Optional)
#     Specify a custom output folder path (relative or absolute) to save the generated shortcuts. If not provided, a folder named "Shell Folder Shortcuts" will be created in the script's directory.
#
# -DeletePreviousOutputFolder
#     Switch (Takes no values)
#     Delete the previous output folder before running the script if one exists matching the one that would be created
#
# -Verbose
#     Switch (Takes no values)
#     Enable verbose output for more detailed information during script execution
#
# -DontGroupTasks
#     Switch (Takes no values)
#     Prevent grouping task shortcuts, meaning the application name won't be prepended to the task name in the shortcut file
#
# -UseAlternativeCategoryNames
#     Switch (Takes no values)
#     Looks up alternative category names for task links to prepend to the task names
#
# -SkipCLSID
#     Switch (Takes no values)
#     Skip creating shortcuts for shell folders based on CLSIDs
#
# -SkipNamedFolders
#     Switch (Takes no values)
#     Skip creating shortcuts for named special folders
#
# -SkipTaskLinks
#     Switch (Takes no values)
#     Skip creating shortcuts for task links (sub-pages within shell folders and control panel menus)
#
# -SkipMSSettings
#     Switch (Takes no values)
#     Skip creating shortcuts for ms-settings: links (system settings pages)
#
# -SkipDeepLinks
#     Switch (Takes no values)
#     Skip creating shortcuts for deep links (links to specific settings pages within system settings)
#
# -SkipURLProtocols
#     Switch (Takes no values)
#     Skip creating shortcuts for URL protocols (e.g., mailto:, ms-settings:, etc.)
#
# -CustomDLLPath
#     String (Optional)
#     Specify a custom DLL file path to load the shell32.dll content from. If not provided, the default shell32.dll will be used.
#     NOTE: Because of how Windows works behind the scenes, DLLs reference resources in corresponding .mui and .mun files.
#        The XML data (resource ID 21) that is required in this script is actually located in shell32.dll.mun, which is located at "C:\Windows\SystemResources\shell32.dll.mun"
#        This means if you want to reference the data from a DLL that is NOT at C:\Windows\System32\shell32.dll, you should directly reference the .mun file instead. It will auto redirect if it's at that exact path, but not otherwise.
#            > This is especially important if wanting to reference the data from a different computer, you need to be sure to copy the .mun file
#        See: https://stackoverflow.com/questions/68389730/windows-dll-function-behaviour-is-different-if-dll-is-moved-to-different-locatio
#
# -CustomLanguageFolderPath
#     String (Optional)
#     Specify a path to a folder containing language-specific MUI files to use for localized string references, and it will prefer any mui files from there if available instead of the system default.
#     For example, to use your own language file for shell32.dll, you could specify a path to a folder containing a file named "shell32.dll.mui" in the desired language, and any other such files.
#     For another example, if you have multiple language packs installed on your system, you could specify the entire language directory in system32 such as "C:\Windows\System32\en-US" to use English strings, or "C:\Windows\System32\de-DE" for German strings.
#
# -CustomSystemSettingsDLLPath
#     String (Optional)
#     Specify a custom path to the SystemSettings.dll file to load the system settings (ms-settings: links) content from. If not provided, the default SystemSettings.dll will be used.
#
# ---------------------------------------------------------------------
#
#   EXAMPLE USAGE FROM COMMAND LINE:
#       .\Super_God_Mode.ps1 -SaveXML -SaveCSV
#
# ---------------------------------------------------------------------
[CmdletBinding()]
param(
    [switch]$DontGroupTasks,
    [switch]$SaveXML,
    [switch]$SaveCSV,
    [switch]$DeletePreviousOutputFolder,
    [string]$CustomDLLPath,
    [string]$CustomLanguageFolderPath,
    [string]$CustomSystemSettingsXMLPath,
    [string]$Output,
    [switch]$SkipCLSID,
    [switch]$SkipNamedFolders,
    [switch]$SkipTaskLinks,
    [switch]$SkipMSSettings,
    [switch]$SkipDeepLinks,
    [switch]$SkipURLProtocols,
    [switch]$UseAlternativeCategoryNames
)

# Set the output folder path for the generated shortcuts based on the provided argument or default location. Convert to full path if necessary
if ($Output) {
    # Convert to full path only if necessary, otherwise use as is
    if (-not [System.IO.Path]::IsPathRooted($Output)) {
        $mainShortcutsFolder = Join-Path $PSScriptRoot $Output
    } else {
        $mainShortcutsFolder = $Output
    }
} else {
    # Default output folder path is a subfolder named "Shell Folder Shortcuts" in the script's directory
    $mainShortcutsFolder = Join-Path $PSScriptRoot "Shell Folder Shortcuts"
}

# Construct paths for subfolders
$CLSIDshortcutsOutputFolder = Join-Path $mainShortcutsFolder "CLSID Shell Folder Shortcuts"
$namedShortcutsOutputFolder = Join-Path $mainShortcutsFolder "Named Shell Folder Shortcuts"
$taskLinksOutputFolder = Join-Path $mainShortcutsFolder "All Task Links"
$msSettingsOutputFolder = Join-Path $mainShortcutsFolder "System Settings"
$deepLinksOutputFolder = Join-Path $mainShortcutsFolder "Deep Links"
$URLProtocolLinksOutputFolder = Join-Path $mainShortcutsFolder "URL Protocols"
$statisticsOutputFolder = Join-Path $mainShortcutsFolder "_Script Result Statistics"

# Set filenames for various output files (CSV and optional XML)
$clsidCsvPath = Join-Path $statisticsOutputFolder "CLSID_Shell_Folders.csv"
$namedFoldersCsvPath = Join-Path $statisticsOutputFolder "Named_Shell_Folders.csv"
$taskLinksCsvPath = Join-Path $statisticsOutputFolder "Task_Links.csv"
$msSettingsCsvPath = Join-Path $statisticsOutputFolder "MS_Settings.csv"
$deepLinksCsvPath = Join-Path $statisticsOutputFolder "Deep_Links.csv"
$URLProtocolLinksCsvPath = Join-Path $statisticsOutputFolder "URL_Protocols.csv"

# XML content file paths
$xmlContentFilePath = Join-Path $statisticsOutputFolder "Shell32_XML_Content.xml"
$resolvedXmlContentFilePath = Join-Path $statisticsOutputFolder "Shell32_XML_Content_Resolved.xml"
$resolvedSettingsXmlContentFilePath = Join-Path $statisticsOutputFolder "Settings_XML_Content_Resolved.xml"
#$msSettingsListFilePath = Join-Path $statisticsOutputFolder "MS_Settings_List.txt"

# Other constants / known paths.
# Available AllSystemSettings XML files may differ depending on Windows 11 or Windows 10, so will try them in order:
$allSettingsXmlPath1 = "C:\Windows\ImmersiveControlPanel\Settings\AllSystemSettings_{D6E2A6C6-627C-44F2-8A5C-4959AC0C2B2D}.xml"
$allSettingsXmlPath2 = "C:\Windows\ImmersiveControlPanel\Settings\AllSystemSettings_{FDB289F3-FCFC-4702-8015-18926E996EC1}.xml"
$allSettingsXmlPath3 = "C:\Windows\ImmersiveControlPanel\Settings\AllSystemSettings_{253E530E-387D-4BC2-959D-E6F86122E5F2}.xml"
$systemSettingsDllPath = "C:\Windows\ImmersiveControlPanel\SystemSettings.dll"

# Check which AllSystemSettings XML file to use
if (Test-Path $allSettingsXmlPath1) {
    $allSettingsXmlPath = $allSettingsXmlPath1
} elseif (Test-Path $allSettingsXmlPath2) {
    $allSettingsXmlPath = $allSettingsXmlPath2
} elseif (Test-Path $allSettingsXmlPath3) {
    $allSettingsXmlPath = $allSettingsXmlPath3
} else {
    Write-Warning "No AllSystemSettings XML file found. Deep Link shortcuts will not be created."
}

# Creates the main directory if it does not exist; `-Force` ensures it is created without errors if it already exists. It won't overwrite any files within even if the folder already exists
try {
    New-Item -Path $mainShortcutsFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
# If creating the folder failed and it doesn't already exist, throw an error and exit the script. Give suggestions for some specific cases
} catch [System.UnauthorizedAccessException] {
    Write-Error "Failed to create output folder: $_"
    # If the default path is used
    if (-not $Output) {
        Write-Error "This may be due to a permissions issue. Ensure you have permissions to create a folder in the script's directory."
    } else {
        Write-Error "This may be due to a permissions issue. Ensure you have permissions to create a folder at the specified path."
    }
    return
} catch {
    if (-not (Test-Path $mainShortcutsFolder)) {
        Write-Error "Failed to create output folder: $_"
        return
    }
}

# If the -DeletePreviousOutputFolder switch is used, go into the set main folder and delete each set subfolder using above variable names
# Doing this instead of just deleting the entire main folder in case the user wants to put the output into a directory in use for other things
if ($DeletePreviousOutputFolder) {
    try {
        if (Test-Path $mainShortcutsFolder) {
            # Remove folders
            if (Test-Path $CLSIDshortcutsOutputFolder){
                Remove-Item -Path $mainShortcutsFolder -Recurse -Force
            }
            if (Test-Path $namedShortcutsOutputFolder){
                Remove-Item -Path $namedShortcutsOutputFolder -Recurse -Force
            }
            if (Test-Path $taskLinksOutputFolder){
                Remove-Item -Path $taskLinksOutputFolder -Recurse -Force
            }
            if (Test-Path $statisticsOutputFolder) {
                Remove-Item -Path $statisticsOutputFolder -Recurse -Force
            }
            if (Test-Path $msSettingsOutputFolder){
                Remove-Item -Path $msSettingsOutputFolder -Recurse -Force
            }
            if (Test-Path $deepLinksOutputFolder){
                Remove-Item -Path $deepLinksOutputFolder -Recurse -Force
            }
            if (Test-Path $URLProtocolLinksOutputFolder){
                Remove-Item -Path $URLProtocolLinksOutputFolder -Recurse -Force
            }
        }
    } catch {
        Write-Error "Failed to delete contents of previous output folder: $_"
    }
}

# Validate the custom dll path if provided
if ($CustomDLLPath) {
    if (-not (Test-Path $CustomDLLPath)) {
        Write-Error "The specified DLL path does not exist: $CustomDLLPath"
        return
    }
}

# Validate the custom language folder path if provided. Ensure it is a folder
if ($CustomLanguageFolderPath) {
    if (-not (Test-Path $CustomLanguageFolderPath -PathType Container)) {
        Write-Error "The specified custom language folder path is not a valid folder: $CustomLanguageFolderPath"
        # Check if they insetad provided a file path, and if so, suggest they provide the folder containing the file
        if (Test-Path $CustomLanguageFolderPath -PathType Leaf) {
            Write-Error "If you are trying to specify a file, please provide the folder containing the file instead, and name it to correspond with whatever DLL file it is for."
        }
        return
    }
    else {
        Write-Verbose "Using custom language folder path: $CustomLanguageFolderPath"
    }
}

# Validate the custom system settings DLL path if provided
if ($CustomSystemSettingsDLLPath) {
    if (-not (Test-Path $CustomSystemSettingsDLLPath)) {
        Write-Error "The specified SystemSettings.dll path does not exist: $CustomSystemSettingsDLLPath"
        return
    } else {
        $systemSettingsDllPath = $CustomSystemSettingsDLLPath
    }
}

# Function to create a folder with a custom icon and set Details view
function New-FolderWithIcon {
    param (
        [string]$FolderPath,
        [string]$IconFile,
        [string]$IconIndex
    )
    # Create the folder
    New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null

    # If there's not a negative sign at the beginning of the index, add one
    if ($IconIndex -notmatch '^-') {
        $IconIndex = "-$IconIndex"
    }

    # Create desktop.ini content
    $desktopIniContent = @"
[.ShellClassInfo]
IconResource=$IconFile,$IconIndex
[ViewState]
Mode=4
Vid={137E7700-3573-11CF-AE69-08002B2E1262}
"@

    # Create desktop.ini file
    $desktopIniPath = Join-Path $FolderPath "desktop.ini"
    Set-Content -Path $desktopIniPath -Value $desktopIniContent -Encoding ASCII -Force

    # Set desktop.ini attributes
    $desktopIniItem = Get-Item $desktopIniPath -Force
    $desktopIniItem.Attributes = 'Hidden,System'

    # Set folder attributes
    $folderItem = Get-Item $FolderPath -Force
    $folderItem.Attributes = 'ReadOnly,Directory'
}


# Create relevant subfolders for different types of shortcuts, and set custom icons for each folder
# Notes for choosing an icon:
#    - You can use the tool 'IconsExtract' from NirSoft to see icons in a DLL file and their indexes: https://www.nirsoft.net/utils/iconsext.html
#    - Another good dll to use for icons is "C:\Windows\System32\imageres.dll" which has a lot of icons

if (-not $SkipCLSID) {
    New-FolderWithIcon -FolderPath $CLSIDshortcutsOutputFolder -IconFile "C:\Windows\System32\shell32.dll" -IconIndex "20"
}

if (-not $SkipNamedFolders) {
    New-FolderWithIcon -FolderPath $namedShortcutsOutputFolder -IconFile "C:\Windows\System32\shell32.dll" -IconIndex "46"
}

if (-not $SkipTaskLinks) {
    New-FolderWithIcon -FolderPath $taskLinksOutputFolder -IconFile "C:\Windows\System32\shell32.dll" -IconIndex "137"
}
if (-not $SkipMSSettings) {
    New-FolderWithIcon -FolderPath $msSettingsOutputFolder -IconFile "C:\Windows\System32\imageres.dll" -IconIndex "114"
}
if (-not $SkipDeepLinks) {
    New-FolderWithIcon -FolderPath $deepLinksOutputFolder -IconFile "C:\Windows\System32\imageres.dll" -IconIndex "114"
}
if (-not $SkipURLProtocols) {
    New-FolderWithIcon -FolderPath $URLProtocolLinksOutputFolder -IconFile "C:\Windows\System32\imageres.dll" -IconIndex "114"
}

# If -SaveCSV or -SaveXML switches are used, create the statistics folder and set to default folder icon
if ($SaveCSV -or $SaveXML) {
    New-FolderWithIcon -FolderPath $statisticsOutputFolder -IconFile "C:\Windows\System32\imageres.dll" -IconIndex "3"
}


# The following block adds necessary .NET types to PowerShell for later use.
# The `Add-Type` cmdlet is used to add C# code that interacts with Windows API functions.
# These functions include loading and freeing DLLs, finding and loading resources, and more.
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    public class Windows
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr FindResource(IntPtr hModule, int lpName, string lpType);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr LockResource(IntPtr hResData);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool FreeLibrary(IntPtr hModule);
    }
"@

# P/Invoke definitions for icon extraction from executables
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class IconExtractor
{
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
"@

# For the Get-LocalizedString function, add a type definition for the Win32 class to load indirect strings from DLLs
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

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool FreeLibrary(IntPtr hModule);
    }
"@
}

# Function: Get-LocalizedString
# This function retrieves a localized (meaning in the user's language) string from a DLL based on a reference string given in the registry
# `StringReference` is a reference in the format "@<dllPath>,-<resourceId>".
function Get-LocalizedString {
    param (
        [string]$StringReference,
        [string]$CustomLanguageFolder
    )

    # Check if it's the special case with multiple concatenated references
    if ($StringReference -match '^\@\@') {
        $references = $StringReference -split '@' | Where-Object { $_ -ne '' }
        $resolvedStrings = @()
        foreach ($ref in $references) {
            $resolved = Get-LocalizedString "@$ref" $CustomLanguageFolder
            if ($resolved) {
                $resolvedStrings += $resolved -split ';' | ForEach-Object { $_.Trim() }
            }
        }
        return ($resolvedStrings | Select-Object -Unique) -join ';'
    }
    # Check if the string is an ms-resource reference
    elseif ($StringReference -match '@\{.+\?ms-resource://.+}') {
        return Get-MsResource $StringReference
    }
    # Existing logic for DLL-based references
    elseif ($StringReference -match '@(.+),-(\d+)') {
        $dllPath = [Environment]::ExpandEnvironmentVariables($Matches[1])
        $resourceId = [uint32]$Matches[2]

        # Custom language folder logic...

        $hModule = [Win32]::LoadLibrary($dllPath)
        if ($hModule -eq [IntPtr]::Zero) {
            Write-Error "Failed to load library: $dllPath"
            return $null
        }

        $stringBuilder = New-Object System.Text.StringBuilder 1024
        $result = [Win32]::LoadString($hModule, $resourceId, $stringBuilder, $stringBuilder.Capacity)

        [void][Win32]::FreeLibrary($hModule)

        if ($result -ne 0) {
            return $stringBuilder.ToString()
        } else {
            Write-Error "Failed to load string resource: $resourceId from $dllPath"
            return $null
        }
    } else {
        Write-Error "Invalid string reference format: $StringReference"
        return $null
    }
}

function Get-MsResource {
    param (
        [string]$ResourcePath
    )
    Write-Verbose "Retrieving resource: $ResourcePath"
    $stringBuilder = New-Object System.Text.StringBuilder 1024
    $result = [Win32]::SHLoadIndirectString($ResourcePath, $stringBuilder, $stringBuilder.Capacity, [IntPtr]::Zero)
    if ($result -eq 0) {
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
        Write-Verbose "Failed to retrieve without /resources/ folder. Error code: $result"

        Write-Error "Failed to retrieve ms-resource: $ResourcePath. Error code: $result"
        return $null
    }
}

# Function: Get-FolderName
# This function retrieves the name of a shell folder given its CLSID, to be used for the shortcuts later
# It attempts to find the name by checking several potential locations in the registry.
function Get-FolderName {
    param (
        [string]$clsid,  # The CLSID of the shell folder.
        [string]$CustomLanguageFolder  # Optional: Path to a folder containing language-specific MUI files for localized string references.
    )

    # Initialize $nameSource to track where the folder name was found (for reporting purposes in CSV later)
    $nameSource = "Unknown"

    Write-Verbose "Attempting to get folder name for CLSID: $clsid"

    # Step 1: Check the default value in the registry at HKEY_CLASSES_ROOT\CLSID\<clsid>.
    $defaultPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid"
    Write-Verbose "Checking default value at: $defaultPath"
    $defaultName = (Get-ItemProperty -Path $defaultPath -ErrorAction SilentlyContinue).'(default)'

    # If a default name is found, check if it's a reference to a localized string.
    if ($defaultName) {
        Write-Verbose "Found default name: $defaultName"
        if ($defaultName -match '@.+,-\d+') {
            Write-Verbose "Default name is a localized string reference"
            $resolvedName = Get-LocalizedString -StringReference $defaultName -CustomLanguageFolder $CustomLanguageFolder
            if ($resolvedName) {
                $nameSource = "Localized String"
                Write-Verbose "Resolved default name to: $resolvedName"
                return @($resolvedName, $nameSource)
            }
            else {
                Write-Verbose "Failed to resolve default name, using original value"
            }
        }
        $nameSource = "Default Value"
        return @($defaultName, $nameSource)
    }
    else {
        Write-Verbose "No default name found"
    }

    # Step 2: Check for a `TargetKnownFolder` in the registry, which points to a known folder.
    $initPropertyBagPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\Instance\InitPropertyBag"
    Write-Verbose "Checking for TargetKnownFolder at: $initPropertyBagPath"
    $targetKnownFolder = (Get-ItemProperty -Path $initPropertyBagPath -ErrorAction SilentlyContinue).TargetKnownFolder

    # If a TargetKnownFolder is found, check its description in the registry.
    if ($targetKnownFolder) {
        Write-Verbose "Found TargetKnownFolder: $targetKnownFolder"
        $folderDescriptionsPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\$targetKnownFolder"
        Write-Verbose "Checking for folder name at: $folderDescriptionsPath"
        $folderName = (Get-ItemProperty -Path $folderDescriptionsPath -ErrorAction SilentlyContinue).Name
        if ($folderName) {
            $nameSource = "Known Folder ID"
            Write-Verbose "Found folder name: $folderName"
            return @($folderName, $nameSource)
        }
        else {
            Write-Verbose "No folder name found in FolderDescriptions"
        }
    }
    else {
        Write-Verbose "No TargetKnownFolder found"
    }

    # Step 3: Check for a `LocalizedString` value in the CLSID registry key.
    $localizedStringPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid"
    Write-Verbose "Checking for LocalizedString at: $localizedStringPath"
    $localizedString = (Get-ItemProperty -Path $localizedStringPath -ErrorAction SilentlyContinue).LocalizedString

    # If a LocalizedString is found, resolve it using the Get-LocalizedString function.
    if ($localizedString) {
        Write-Verbose "Found LocalizedString: $localizedString"
        $resolvedString = Get-LocalizedString -StringReference $localizedString -CustomLanguageFolder $CustomLanguageFolder
        if ($resolvedString) {
            $nameSource = "Localized String"
            Write-Verbose "Resolved LocalizedString to: $resolvedString"
            return @($resolvedString, $nameSource)
        }
        else {
            Write-Verbose "Failed to resolve LocalizedString"
        }
    }
    else {
        Write-Verbose "No LocalizedString found"
    }

    # Step 4: Check the Desktop\NameSpace registry key for the folder name.
    $namespacePath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$clsid"
    Write-Verbose "Checking Desktop\NameSpace at: $namespacePath"
    $namespaceName = (Get-ItemProperty -Path $namespacePath -ErrorAction SilentlyContinue).'(default)'

    # If a name is found here, return it.
    if ($namespaceName) {
        $nameSource = "Desktop Namespace"
        Write-Verbose "Found name in Desktop\NameSpace: $namespaceName"
        return @($namespaceName, $nameSource)
    }
    else {
        Write-Verbose "No name found in Desktop\NameSpace"
    }

    # Step 5: New check - Recursively search in HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer
    $explorerPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
    Write-Verbose "Recursively checking Explorer registry path for CLSID: $explorerPath"

    function Search-RegistryKey {
        param (
            [string]$Path,
            [string]$Clsid
        )

        $keys = Get-ChildItem -Path $Path -ErrorAction SilentlyContinue

        foreach ($key in $keys) {
            if ($key.PSChildName -eq $Clsid) {
                $defaultValue = (Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue).'(default)'
                if ($defaultValue -and $defaultValue -ne "") {
                    return $defaultValue
                }
            }

            $subResult = Search-RegistryKey -Path $key.PSPath -Clsid $Clsid
            if ($subResult) {
                return $subResult
            }
        }

        return $null
    }

    $explorerName = Search-RegistryKey -Path $explorerPath -Clsid $clsid
    if ($explorerName) {
        $nameSource = "Explorer Registry"
        Write-Verbose "Found name in Explorer registry: $explorerName"
        return @($explorerName, $nameSource)
    }
    else {
        Write-Verbose "No name found in Explorer registry"
    }

    # Step 6: If no name is found, return the CLSID itself as a last resort to be used for the shortcut
    Write-Verbose "Returning CLSID as folder name"
    $nameSource = "Unknown"
    return @($clsid, $nameSource)
}

# Function: Create-Shortcut
# This function creates a shortcut for a given CLSID or shell folder.
# It assigns the appropriate target path, arguments, and icon based on the CLSID information.
function Create-Shortcut {
    param (
        [string]$clsid,         # The CLSID of the shell folder
        [string]$name,          # The name of the shortcut
        [string]$shortcutPath,  # The full path where the shortcut will be created
        [string]$pageName = ""  # Optional: the name of a specific page within the shell folder (usually used for control panels)
    )

    try {
        Write-Verbose "Creating Shortcut For: $name"
        # Create a COM object representing the WScript.Shell, which is used to create shortcuts
        $shell = New-Object -ComObject WScript.Shell

        # Create the actual shortcut at the specified path
        $shortcut = $shell.CreateShortcut($shortcutPath)
        # Set the 'target' to explorer so it opens with File Explorer. The 'arguments' part of the target will be set next and contain the 'shell:' part of the command.
        $shortcut.TargetPath = "explorer.exe"

        # If a specific page is provided, include it in the arguments, otherwise just set the shell command used to open the folder
        if ($pageName) {
            $shortcut.Arguments = "shell:::$clsid\$pageName"
        } else {
            $shortcut.Arguments = "shell:::$clsid"
        }

        # Attempt to find a custom icon for the shortcut by checking the registry
        $iconPath = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\DefaultIcon" -ErrorAction SilentlyContinue).'(default)'
        if ($iconPath) {
            Write-Verbose "Setting custom icon: $iconPath"
            $shortcut.IconLocation = $iconPath
        }
        # Otherwise use the Windows default folder icon
        else {
            Write-Verbose "No custom icon found. Setting default folder icon."
            $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,3"
        }

        $shortcut.Save()

        # Release the COM object to free up resources.
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $true
    }
    catch {
        Write-Host "Error creating shortcut for $name`: $($_.Exception.Message)"
        return $false
    }
}

function Get-TaskIcon {
    param (
        [string]$controlPanelName,
        [string]$applicationId,
        [string]$commandTarget
    )

    $iconPath = $null

    if ($controlPanelName) {
        # Try to get icon from control panel name
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\$controlPanelName"
        $iconPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Icon
    }

    if (-not $iconPath -and $applicationId) {
        Write-Verbose "No icon found for control panel name, trying registry via application ID: $applicationId"
        # Try to get icon from application ID
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\$applicationId"
        $iconPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Icon

        if (-not $iconPath) {
            # If not found, try CLSID path
            $regPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$applicationId\DefaultIcon"
            $iconPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).'(default)'
        }
    }

    # If icon path still not found but there is a command, see if the command calls an exe and use that exe's icon if so. Except if it's control.exe
    if (-not $iconPath -and $commandTarget) {
        Write-Verbose "No icon found in registry. Checking if icon can be extracted from command target: $commandTarget"
        # Extract the file name from the path if it's not just the filename
        $commandFileName = Split-Path -Path $commandTarget -Leaf
        # Check if it's an exe. Also don't use icons from a few specific executables like control.exe
        $ignoredExeFiles = @("control.exe", "rundll32.exe")
        if ($commandFileName -match '\.exe$' -and $ignoredExeFiles -notcontains $commandFileName) {
            try {
                $hIcon = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $commandTarget, 0)
                if ($hIcon -ne [IntPtr]::Zero) {
                    Write-Verbose "Using embedded icon from $commandTarget"
                    $iconPath = $commandTarget + ",0"
                    [void][IconExtractor]::DestroyIcon($hIcon)
                } else {
                    Write-Verbose "No embedded icon found in $commandTarget"
                }
            } catch {
                Write-Verbose "Failed to extract icon from $commandTarget`: $_"
            }
        }
        else {
            Write-Verbose "Command target is not an exe or is part of ignored exe list - Ignoring."
        }
    }

    if ($iconPath) {
        return Fix-CommandPath $iconPath
    }

    # Default icon if none found
    Write-Verbose "Using Default Icon - No available icon found for $controlPanelName or $applicationId"
    return "%SystemRoot%\System32\shell32.dll,0"
}

# Fix issues with command paths spotted in at least one XML entry, where double percent signs were present at the beginning like %%windir%
function Fix-CommandPath {
    param (
        [string]$command
    )

    # Fix double % at the beginning
    if ($command -match '^\%\%') {
        $command = $command -replace '^\%\%', '%'
    }

    # Expand environment variables
    #$command = [Environment]::ExpandEnvironmentVariables($command)

    return $command
}

function Create-TaskLink-Shortcut {
    param (
        [string]$name,
        [string]$shortcutPath,
        [string]$shortcutType,
        [string]$command,
        [string]$controlPanelName,
        [string]$applicationId,
        [string[]]$keywords
    )

    try {
        Write-Verbose "Creating Task Link Shortcut For: $name"

        $shell = New-Object -ComObject WScript.Shell

        if ($shortcutType -eq "url") {
            # For URL shortcuts
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $command
        } else {
            # For regular shortcuts
            $shortcut = $shell.CreateShortcut($shortcutPath)

            # Parse the command
            if ($command -match '^(\S+)\s*(.*)$') {
                $targetPath = Fix-CommandPath $Matches[1]
                $arguments = Fix-CommandPath $Matches[2]

                # Expand variables in the arguments such as %windir%, because shortcuts don't seem to work with them in the arguments
                $arguments = [Environment]::ExpandEnvironmentVariables($arguments)

                $shortcut.TargetPath = $targetPath
                $shortcut.Arguments = $arguments
            } else {
                $fixedCommand = Fix-CommandPath $command
                $shortcut.TargetPath = $fixedCommand
            }

            # Add keywords only if it's a .lnk type shortcut
            # Combine keywords into a single string and set as Description
            if ($keywords -and $keywords.Count -gt 0) {
                $descriptionLimit = 259 # Limit for Description field in shortcuts or else it causes some kind of buffer overflow
                $keywordString = ""
                $separator = " "

                foreach ($keyword in $keywords) {
                    $potentialNewString = if ($keywordString) { $keywordString + $separator + $keyword } else { $keyword }
                    if ($potentialNewString.Length -le $descriptionLimit) {
                        $keywordString = $potentialNewString
                    } else {
                        break
                    }
                }

                $shortcut.Description = $keywordString
            }
        }

        $iconPath = Get-TaskIcon -controlPanelName $controlPanelName -applicationId $applicationId -commandTarget $shortcut.TargetPath

        # Only add icon at this point if it's a .lnk type shortcut, not for .url which needs to be done after
        if ($shortcutType -eq "lnk") {
            $shortcut.IconLocation = $iconPath
        }
        $shortcut.Save()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null

        # For .url type shortcuts, open the plaintext url file and add the IconIndex= and IconFile= lines to the end
        if ($shortcutType -eq "url") {
            if ($iconPath) {
                $iconFile, $iconIndex = $iconPath -split ','
                # Append the icon information to the file
                Add-Content -Path $shortcutPath -Value "IconFile=$iconFile"
                if ($iconIndex) {
                    Add-Content -Path $shortcutPath -Value "IconIndex=$iconIndex"
                }
            }
        }

        return $true
    } catch {
        Write-Host "Error creating shortcut for $name`: $($_.Exception.Message)"
        return $false
    }
}

# Function: Get-Shell32XMLContent
# This function extracts and returns the XML content embedded in the shell32.dll file, which contains info about sub-pages within certain shell folders and control panel menus.
# Apparently these are sometimes referred to as "Task Links"
function Get-Shell32XMLContent {
    param(
        [switch]$SaveXML,
        [string]$CustomDLL
    )

    # If a custom DLL path is provided, use it; otherwise, use the default shell32.dll path
    if ($CustomDLL) {
        Write-Verbose "Using custom DLL path: $CustomDLL"
        $dllPath = $CustomDLL
    } else {
        Write-Verbose "Using default shell32.dll path"
        $dllPath = "shell32.dll"
        # Join to system path
        $dllPath = Join-Path $env:SystemRoot "System32\$dllPath"
    }

    # Constants used for loading the shell32.dll as a data file.
    $LOAD_LIBRARY_AS_DATAFILE = 0x00000002
    $DONT_RESOLVE_DLL_REFERENCES = 0x00000001

    # Initialize an empty string to hold the XML content
    $xmlContent = ""

    # Load shell32.dll as a data file, preventing the DLL from being fully resolved as it is not necessary
    $shell32Handle = [Windows]::LoadLibraryEx($dllPath, [IntPtr]::Zero, $LOAD_LIBRARY_AS_DATAFILE -bor $DONT_RESOLVE_DLL_REFERENCES)
    if ($shell32Handle -eq [IntPtr]::Zero) {
        Write-Error "Failed to load $dllPath"
        return $null
    }

    try {
        # Attempt to find the XML resource within shell32.dll. '21' is necessary to use here.
        $hResInfo = [Windows]::FindResource($shell32Handle, 21, "XML")
        if ($hResInfo -eq [IntPtr]::Zero) {
            Write-Error "Failed to find XML resource"
            Write-Error "Did you move the DLL from the original location? If so you may need to directly specify the corresponding .mun file instead of the DLL."
            Write-Error "See the comments for the DLLPath argument at the top of the script for more info."
            return $null
        }

        # Load the XML resource data.
        $hResData = [Windows]::LoadResource($shell32Handle, $hResInfo)
        if ($hResData -eq [IntPtr]::Zero) {
            Write-Error "Failed to load XML resource"
            return $null
        }

        # Lock the resource in memory to access its data.
        $pData = [Windows]::LockResource($hResData)
        if ($pData -eq [IntPtr]::Zero) {
            Write-Error "Failed to lock XML resource"
            return $null
        }

        # Get the size of the XML resource and copy it into a byte array.
        $size = [Windows]::SizeofResource($shell32Handle, $hResInfo)
        $byteArray = New-Object byte[] $size
        [System.Runtime.InteropServices.Marshal]::Copy($pData, $byteArray, 0, $size)

        # Convert the byte array to a UTF-8 string.
        $xmlContent = [System.Text.Encoding]::UTF8.GetString($byteArray)
        $xmlContent = $xmlContent -replace "`0", "" # Remove any null characters from the string, though this probably isn't necessary
    }
    finally {
        # Ensure that the loaded DLL is always freed from memory.
        [void][Windows]::FreeLibrary($shell32Handle)
    }

    # Clean and trim any extraneous whitespace from the XML content
    $xmlContent = $xmlContent.Trim()

    # Save XML content if the SaveXML switch is used
    if ($SaveXML) {
        Save-PrettyXML -xmlContent $xmlContent -outputPath $xmlContentFilePath
    }

    # Return the XML content as a string
    return $xmlContent
}

# Save the XML content from shell32.dll to a file for reference if the user uses the -SaveXML switch
function Save-PrettyXML {
    param (
        [string]$xmlContent,
        [string]$outputPath
    )

    try {
        # Load the XML content
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.LoadXml($xmlContent)

        # Create XmlWriterSettings for pretty-printing
        $writerSettings = New-Object System.Xml.XmlWriterSettings
        $writerSettings.Indent = $true
        $writerSettings.IndentChars = "  "
        $writerSettings.NewLineChars = "`r`n"
        $writerSettings.NewLineHandling = [System.Xml.NewLineHandling]::Replace

        # Create XmlWriter and write the document
        $writer = [System.Xml.XmlWriter]::Create($outputPath, $writerSettings)
        $xmlDoc.Save($writer)
        $writer.Close()

        Write-Verbose "XML content formatted and saved: $outputPath"
    }
    catch {
        Write-Error "Failed to format and save XML: $_"
    }
}

# Function: Get-TaskLinks
# This function parses the XML content extracted from shell32.dll to find "task links", which are basically sub-menu pages, often in the Control Panel
function Get-TaskLinks {
    param(
        [switch]$SaveXML,
        [string]$DLLPath,
        [string]$CustomLanguageFolder
    )
    $xmlContent = Get-Shell32XMLContent -SaveXML:$SaveXML -CustomDLL:$DLLPath

    try {
        $xml = [xml]$xmlContent
        Write-Verbose "XML parsed successfully."
    } catch {
        Write-Error "Failed to parse XML content: $_"
        return
    }

    # Create a copy of the XML for resolved content
    $resolvedXml = $xml.Clone()

    $nsManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsManager.AddNamespace("cpl", "http://schemas.microsoft.com/windows/cpltasks/v1")
    $nsManager.AddNamespace("sh", "http://schemas.microsoft.com/windows/tasks/v1")
    $nsManager.AddNamespace("sh2", "http://schemas.microsoft.com/windows/tasks/v2")

    $tasks = @()

    $allTasks = $xml.SelectNodes("//sh:task", $nsManager)

    foreach ($task in $allTasks) {
        $taskId = $task.GetAttribute("id")
        $nameNode = $task.SelectSingleNode("sh:name", $nsManager)
        $controlPanel = $task.SelectSingleNode("sh2:controlpanel", $nsManager)
        $commandNode = $task.SelectSingleNode("sh:command", $nsManager)
        $keywordsNodes = $task.SelectNodes("sh:keywords", $nsManager)

        # Resolve name
        $name = $null
        if ($nameNode -and $nameNode.InnerText) {
            if ($nameNode.InnerText -match '@(.+),-(\d+)') {
                $name = Get-LocalizedString -StringReference $nameNode.InnerText -CustomLanguageFolder $CustomLanguageFolder
                # Update resolved XML
                $resolvedNameNode = $resolvedXml.SelectSingleNode("//sh:task[@id='$taskId']/sh:name", $nsManager)
                if ($resolvedNameNode) {
                    $resolvedNameNode.InnerText = $name
                }
            } else {
                $name = $nameNode.InnerText
            }
        }
        if ($name) {
            $name = $name.Trim()
        } elseif ($task.Name -eq "sh:task" -and $task.GetAttribute("idref")) {
            Write-Verbose "Skipping category entry: $($task.OuterXml)"
            continue
        } else {
            Write-Warning "Task $taskId is missing a name and is not a category reference. This may indicate an issue: $($task.OuterXml)"
            continue
        }

        $command = $null
        $appName = $null
        $page = $null

        $appId = $task.ParentNode.id
        if (-not $appId) {
            $appId = $task.GetAttribute("id")
        }

        if ($controlPanel) {
            $appName = $controlPanel.GetAttribute("name")
            $page = $controlPanel.GetAttribute("page")
            $command = "control.exe /name $appName"
            if ($page) {
                $command += " /page $page"
            }
        } elseif ($commandNode) {
            $command = $commandNode.InnerText
        }

        $keywords = @()
        foreach ($keywordNode in $keywordsNodes) {
            $keyword = $null
            if ($keywordNode.InnerText -match '@(.+),-(\d+)') {
                $keyword = Get-LocalizedString -StringReference $keywordNode.InnerText -CustomLanguageFolder $CustomLanguageFolder
                # Update resolved XML
                $resolvedKeywordNode = $resolvedXml.SelectSingleNode("//sh:task[@id='$taskId']/sh:keywords[text()='$($keywordNode.InnerText)']", $nsManager)
                if ($resolvedKeywordNode) {
                    $resolvedKeywordNode.InnerText = $keyword
                }
            } else {
                $keyword = $keywordNode.InnerText
            }
            if ($keyword) {
                $splitKeywords = $keyword.Split(';', [StringSplitOptions]::RemoveEmptyEntries)
                foreach ($splitKeyword in $splitKeywords) {
                    if ($splitKeyword) {
                        $keywords += $splitKeyword.Trim()
                    }
                }
            }
        }

        # If no app name, look it up via CLSID in the registry
        if (-not $appName) {
            $appName = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$appId" -ErrorAction SilentlyContinue)."System.ApplicationName"
        }
        # If still no app name, check if any other task has the same app ID with a name, and if so use that, but only the first instance
        if (-not $appName) {
            $otherTask = $tasks | Where-Object { $_.ApplicationId -eq $appId -and $_.ApplicationName } | Select-Object -First 1
            if ($otherTask) {
                $appName = $otherTask.ApplicationName
            }
        }

        if ($name -and ($command -or $appName)) {
            # Determine the ControlPanelName value before creating the object
            if ($controlPanel) {
                $controlPanelName = $controlPanel.GetAttribute("name")
            } else {
                $controlPanelName = $null
            }

            # Now create the $newTask object using the pre-determined $controlPanelName
            $newTask = [PSCustomObject]@{
                TaskId = $taskId
                ApplicationId = $appId
                Name = $name
                ApplicationName = $appName
                Page = $page
                Command = $command
                Keywords = $keywords
                ControlPanelName = $controlPanelName
            }

            # Check if a task with the same name and command already exists
            $isDuplicate = $tasks | Where-Object { $_.Name -eq $newTask.Name -and $_.Command -eq $newTask.Command }

            if (-not $isDuplicate) {
                $tasks += $newTask
            } else {
                Write-Verbose "Skipping duplicate task: $($newTask.Name)"
            }
        }
    }

    # Store the resolved XML content in a new variable
    $resolvedXmlContent = $resolvedXml.OuterXml

    # Save XML content if the SaveXML switch is used
    if ($SaveXML) {
        Save-PrettyXML -xmlContent $resolvedXmlContent -outputPath $resolvedXmlContentFilePath
    }

    return $tasks
}

# Function: Create-NamedShortcut
# This function creates a shortcut for a named special folder.
# The shortcut points directly to the folder using its name (e.g., "Documents", "Pictures").
function Create-NamedShortcut {
    param (
        [string]$name,          # The name of the special folder.
        [string]$shortcutPath,  # The full path where the shortcut will be created.
        [string]$iconPath       # The path to the folder's custom icon (if any).
    )

    try {
        Write-Verbose "Creating named shortcut for $name"
        # Create a COM object representing the WScript.Shell, which is used to create shortcuts.
        $shell = New-Object -ComObject WScript.Shell

        # Create the actual shortcut at the specified path.
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "explorer.exe"  # The shortcut will open in Windows Explorer.
        $shortcut.Arguments = "shell:$name"  # Set the shortcut to open the specified folder by name. This will also be in the target path box for the shortcut.

        # Set the custom icon if one is provided in the registry
        if ($iconPath) {
            Write-Verbose "Setting custom icon: $iconPath"
            $shortcut.IconLocation = $iconPath
        }
        else {
            Write-Verbose "Setting default folder icon"
            $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,3"  # Default folder icon.
        }

        # Save the shortcut.
        $shortcut.Save()

        # Release the COM object to free up resources.
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $true
    }
    catch {
        Write-Host "Error creating shortcut for $name`: $($_.Exception.Message)"
        return $false
    }
}

# Function: Create-CLSIDCsvFile
# This function generates a CSV file containing details about all processed CLSID shell folders.
function Create-CLSIDCsvFile {
    param (
        [string]$outputPath,  # The full path where the CSV file will be saved.
        [array]$clsidData     # An array of objects containing CLSID data.
    )

    # Initialize the CSV content with headers.
    $csvContent = "CLSID,ExplorerCommand,Name,NameSource,CustomIcon`n"

    # Loop through each CLSID data object and append its details to the CSV content.
    foreach ($item in $clsidData) {
        $explorerCommand = "explorer shell:::$($item.CLSID)"  # The command to open the shell folder.
        $iconPath = if ($item.IconPath) {
            "`"$($item.IconPath -replace '"', '""')`""  # Escape double quotes in the icon path.
        } else {
            "None"
        }

        # Escape any double quotes in the name.
        $escapedName = $item.Name -replace '"', '""'

        # Convert the sub-items array to a string, separating items with a pipe character.
        #$subItemsString = ($item.SubItems | ForEach-Object { "$($_.Name):$($_.Page)" }) -join '|'

        # Append the CLSID details to the CSV content.
        $csvContent += "$($item.CLSID),`"$explorerCommand`",`"$escapedName`",$($item.NameSource),$iconPath`n"
    }

    # Write the CSV content to the specified output file.
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

# Function: Create-NamedFoldersCsvFile
# This function generates a CSV file containing details about all processed named special folders.
function Create-NamedFoldersCsvFile {
    param (
        [string]$outputPath  # The full path where the CSV file will be saved.
    )

    # Initialize the CSV content with headers.
    $csvContent = "Name,ExplorerCommand,RelativePath,ParentFolder`n"

    # Retrieve all named special folders from the registry.
    $namedFolders = Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions"

    # Loop through each named folder and append its details to the CSV content.
    foreach ($folder in $namedFolders) {
        $folderProperties = Get-ItemProperty -Path $folder.PSPath
        $name = $folderProperties.Name  # Extract the name of the folder.
        if ($name) {
            $explorerCommand = "explorer shell:$name"  # The command to open the folder.
            $relativePath = $folderProperties.RelativePath -replace ',', '","'  # Escape commas in the relative path.
            $parentFolderGuid = $folderProperties.ParentFolder  # Extract the parent folder GUID (if any).
            $parentFolderName = "None"  # Default value if there is no parent folder.
            if ($parentFolderGuid) {
                # If a parent folder GUID is found, retrieve the name of the parent folder.
                $parentFolderPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\$parentFolderGuid"
                $parentFolderName = (Get-ItemProperty -Path $parentFolderPath -ErrorAction SilentlyContinue).Name
            }

            # Append the named folder details to the CSV content.
            $csvContent += "`"$name`",`"$explorerCommand`",`"$relativePath`",`"$parentFolderName`"`n"
        }
    }

    # Write the CSV content to the specified output file.
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

# Function: Create-TaskLinksCsvFile
# This function generates a CSV file containing details about all processed 'task links' aka sub-pages.
function Create-TaskLinksCsvFile {
    param (
        [string]$outputPath,
        [array]$taskLinksData
    )

    $csvContent = "XMLTaskId,ApplicationId,ApplicationName,Name,Page,Command,Keywords`n"

    foreach ($item in $taskLinksData) {
        $taskId = $item.TaskId -replace '"', '""'
        $applicationId = $item.ApplicationId -replace '"', '""'
        $applicationName = $item.ApplicationName -replace '"', '""'
        $name = $item.Name -replace '"', '""'
        $page = $item.Page -replace '"', '""'
        $command = $item.Command -replace '"', '""'
        $keywords = ($item.Keywords -join ';') -replace '"', '""'

        $csvContent += "`"$taskId`",`"$applicationId`",`"$applicationName`",`"$name`",`"$page`",`"$command`",`"$keywords`"`n"
    }

    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Create-MSSettingsCsvFile {
    param (
        [string]$outputPath,
        [array]$msSettingsList
    )

    $csvContent = "Setting Name,Full Setting Command`n"

    foreach ($fullLink in $msSettingsList) {
        # Split on the first colon to separate the command from the name
        $fullLinkParts = $fullLink -split ":", 2
        $name = $fullLinkParts[1].Trim()

        $csvContent += "`"$name`",`"$fullLink`"`n"
    }

    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

# Take the app name like Microsoft.NetworkAndSharingCenter and prepare it to be displayed in the shortcut name like "Network and Sharing Center - Whatever Task name"
function Prettify-App-Name {
    param(
        [string]$AppName,
        [string]$TaskName
    )

    # List of words to rejoin if split
    $wordsToRejoin = @(
        "Bit Locker",
        "Side Show",
        "Auto Play"
        # Add more words as needed
    )

    # Remove "Microsoft." prefix if present
    $AppName = $AppName -replace '^Microsoft\.', ''

    # Split camelCase into separate words, handling consecutive uppercase letters
    $AppName = $AppName -creplace '(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])|\b(?=[A-Z]{2,}\b)', ' '

    # Rejoin specific words
    foreach ($word in $wordsToRejoin) {
        $AppName = $AppName -replace $word, $word.Replace(' ', '')
    }

    # Combine AppName and TaskName
    $PrettyName = "$AppName - $TaskName"

    # Sanitize the name to remove invalid characters for file names
    $PrettyName = $PrettyName -replace '[\\/:*?"<>|]', '_'

    return $PrettyName
}

# Function to parse the ms-settings XML file and extract relevant data
function Get-AllSettings-Data {
    param (
        [string]$xmlFilePath,
        [switch]$SaveXML
    )

    if (-not (Test-Path $xmlFilePath)) {
        Write-Error "All Systems XML file not found: $xmlFilePath"
        return $null
    }

    try {
        [xml]$xmlContent = Get-Content $xmlFilePath
        $settingsData = @()

        foreach ($content in $xmlContent.PCSettings.SearchableContent) {
            $settingInfo = @{
                Name = $content.Filename
                PageID = $null
                GroupID = $null
                Description = Get-LocalizedString $content.SettingInformation.Description
                HighKeywords = $null
                Glyph = $content.ApplicationInformation.Glyph
                DeepLink = $content.ApplicationInformation.DeepLink
                IconPath = $content.ApplicationInformation.Icon
                PolicyIds = @()
            }

            # Resolve HighKeywords if it exists
            if ($content.SettingInformation.HighKeywords) {
                $settingInfo.HighKeywords = Get-LocalizedString $content.SettingInformation.HighKeywords
            }

            # Extract PageID, GroupID, and PolicyIds from SettingPaths
            $settingPaths = $content.SettingIdentity.SettingPaths.Path
            if ($settingPaths) {
                $firstPath = $settingPaths[0]
                $settingInfo.PageID = $firstPath.PageID
                $settingInfo.GroupID = $firstPath.GroupID

                if ($firstPath.PolicyIds) {
                    $settingInfo.PolicyIds = $firstPath.PolicyIds -split ';' | ForEach-Object { $_.Trim() }
                }
            }

            # Update the XML content with resolved strings
            $content.SettingInformation.Description = $settingInfo.Description
            if ($settingInfo.HighKeywords) {
                $content.SettingInformation.HighKeywords = $settingInfo.HighKeywords
            }

            # Only add to settingsData if it has a policy ID, which is what the ms-settings shortcuts use
            # if ($settingInfo.PolicyIds) {
            #     $settingsData += $settingInfo
            # }

            # Add all data and handle it elsewhere
            $settingsData += $settingInfo
        }

        # Save the resolved XML to the main output folder if the SaveXML switch is used
        if ($SaveXML) {
            $xmlContent.Save($resolvedSettingsXmlContentFilePath)
            Write-Verbose "Resolved XML saved to: $resolvedSettingsXmlContentFilePath"
        }

        return $settingsData
    }
    catch {
        Write-Error "Error parsing MS Settings XML: $_"
        return $null
    }
}

# Function to extract ms-settings links from the SystemSettings.dll file which has the ms-settings links embedded in it
function Get-MS-SettingsFrom-SystemSettingsDLL {
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

    $matchesList = [regex]::Matches($content, 'ms-settings:[a-z-]+')
    foreach ($match in $matchesList) {
        [void]$results.Add($match.Value)
    }

    Write-Verbose "Unique MS-Settings Matches Found: $($results.Count)"
    return $results | Sort-Object
}

# Function to create ms-settings shortcuts
function Create-MSSettings-Shortcut {
    param (
        [string]$fullName,
        [string]$shortcutPath
    )

    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $fullName

        # Use a default icon of settings gear for each ms-settings shortcut
        $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,-16826"

        $shortcut.Save()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $true
    }
    catch {
        Write-Host "Error creating shortcut for $name`: $($_.Exception.Message)"
        return $false
    }
}

function Create-Deep-Link-Shortcut {
    param (
        [object]$settingArray
    )
    $rawTarget = $settingArray.DeepLink

    # If there's a description, use that as the name
    if ($settingArray.Description) {
        $name = $settingArray.Description
    } else {
        # Sanitize the name and use that
        $name = $rawTarget
    }

    # Determine type of link/command. First check if it matches application ID format like "Microsoft.Recovery"
    if ($rawTarget -match '^Microsoft\.[a-zA-Z]+$') {
        $shortcutType = "app"
        $target = "control.exe"
        $targetArgs = "/name $rawTarget"
        $fullCommand = "control.exe /name $rawTarget"
    # Check if it's an application name but with a backslash and therefore has a page like Microsoft.Mouse\2 or Microsoft.PowerOptions\pagePlanSettings
    } elseif ($rawTarget -match '^Microsoft\.[a-zA-Z]+\\[a-zA-Z0-9]+$') {
        $shortcutType = "appPage"
        $target = "control.exe"
        $targetArgs = "/name $($rawTarget.Split('\')[0]) /page $($rawTarget.Split('\')[1])"
        $fullCommand = "$target $targetArgs"
    # Check if it's a shell:::{CLSID} link. It may have stuff after it as part of the link
    } elseif ($rawTarget -match '^shell:::{[a-zA-Z0-9-]+}') {
        $shortcutType = "clsid"
        $fullCommand = "explorer $rawTarget"
    # If it starts with %windir% or %%windir% assume it's a full path to an executable or URL with or without arguments like %windir%\something
    } elseif ($rawTarget -match '^%') {
        $shortcutType = "pathcommand"
        $fullCommand = $rawTarget

        #Split on the first space, set args to the 2nd match, otherwise assume no args
        if ($fullCommand -match '^(\S+)\s*(.*)$') {
            $target = $Matches[1]
            $targetArgs = $Matches[2]
        } else {
            $target = $fullCommand
            $targetArgs = ""
        }
    # If it's just letters and numbers and the <Filename> property starts with "Defender_" then it is a Windows Defender setting whcih is apparently a special case in this file
    } elseif ($rawTarget -match '^[a-zA-Z0-9]+$' -and $settingArray.Name -match '^Defender_') {
        $shortcutType = "windowsdefender"
        $fullCommand = "windowsdefender://$rawTarget"
        $target = "windowsdefender://$rawTarget"
        # Prepend "Windows Defender" to the name
        $name = "Windows Defender - $name"
    # If it's just letters deal with it later, do nothing
    } elseif ($rawTarget -match '^[a-zA-Z]+$') {
        $shortcutType = "unknown"
    # Assume it's a full path to an executable or URL with or without arguments like %windir%\something
    } else {
        $shortcutType = "assumedPath"
        # Try to split on the first space, set args to the 2nd match, otherwise assume no args
        if ($rawTarget -match '^(\S+)\s*(.*)$') {
            $target = $Matches[1]
            $targetArgs = $Matches[2]
        } else {
            $target = $rawTarget
        }
    }

    # Sanitize the name to make it a valid filename
    $sanitizedName = $name -replace '[\\/:*?"<>|]', '_'
    $shortcutPath = Join-Path $deepLinksOutputFolder "$sanitizedName.lnk"
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)

        # Set the target path and arguments based on the type of link
        if ($shortcutType -eq "app" -or $shortcutType -eq "appPage") {
            $shortcut.TargetPath = $target
            $shortcut.Arguments = $targetArgs
        } elseif ($shortcutType -eq "clsid") {
            $shortcut.TargetPath = "explorer.exe"
            $shortcut.Arguments = $rawTarget
        } elseif ($shortcutType -eq "pathcommand") {
            $shortcut.TargetPath = $target
            $shortcut.Arguments = $targetArgs
        } elseif ($shortcutType -eq "assumedPath") {
            $shortcut.TargetPath = $target
        } elseif ($shortcutType -eq "windowsdefender") {
            $shortcut.TargetPath = $target
        } else {
            $shortcut.TargetPath = $rawTarget
        }

        # If there's an icon property try using that
        if ($settingArray.IconPath) {
            $shortcut.IconLocation = $settingArray.IconPath
        } else {
            # Use a default icon of settings gear for each ms-settings shortcut
            $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,-16826"
        }

        $shortcut.Save()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null

        #Add info about the full command and icon path to the settings array and return it updated
        $settingArray.FullCommand = $fullCommand
        $settingArray.ShortcutPath = $shortcutPath
        return $settingArray
    } catch {
        Write-Host "Error creating shortcut for $name`: $($_.Exception.Message)"
        return $false
    }
}

function Create-Deep-Link-CSVFile {
    param (
        [string]$outputPath,
        [array]$deepLinksDataArray
    )

    $csvContent = "Description,Deep Link,Full Command,IconPath`n"

    foreach ($item in $deepLinksDataArray) {
        $description = $item.Description -replace '"', '""'
        $deepLink = $item.DeepLink -replace '"', '""'
        $fullCommand = $item.FullCommand -replace '"', '""'
        $iconPath = if ($item.IconPath) {
            "`"$($item.IconPath -replace '"', '""')`""  # Escape double quotes in the icon path.
        } else {
            "None"
        }

        $csvContent += "`"$description`",`"$deepLink`",`"$fullCommand`",$iconPath`n"
    }

    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Get-And-Process-URL-Protocols {
    # Create object to store data. Will want to store multiple properties for each URL protocol
    $urlProtocols = @()
    $urlProtocolData = @()

    Write-Verbose "Gathering URL Protocol data from the registry"
    $urlProtocols = @{}

    # Function store registry data structure for each protocol in an object so we don't need to make a bunch of registry calls later
    function Get-RegistryKeyData {
        param (
            [Microsoft.Win32.RegistryKey]$Key
        )
        $data = @{
            '(Default)' = $Key.GetValue('')
            Values = @{}
            SubKeys = @{}
        }
        foreach ($valueName in $Key.GetValueNames()) {
            if ($valueName -ne '') {
                $data.Values[$valueName] = $Key.GetValue($valueName)
            }
        }
        foreach ($subKeyName in $Key.GetSubKeyNames()) {
            $subKey = $Key.OpenSubKey($subKeyName)
            $data.SubKeys[$subKeyName] = Get-RegistryKeyData -Key $subKey
            $subKey.Close()
        }
        return $data
    }

    Get-ChildItem -Path 'Registry::HKEY_CLASSES_ROOT' -ErrorAction SilentlyContinue | 
        Where-Object {
            $_.GetValue('(Default)') -match '^URL:' -or $null -ne $_.GetValue('URL Protocol')
        } | ForEach-Object {
            $urlProtocols[$_.PSChildName] = Get-RegistryKeyData -Key $_.OpenSubKey('')
        }

        foreach ($protocol in $urlProtocols.Keys) {
            Write-Verbose "Processing URL Protocol: $protocol"

            $protocolData = $urlProtocols[$protocol]
            $protocolName = $protocolData['(Default)']

            # If the protocol name is in the format "URL:WhateverName", extract the "WhateverName" part
            if ($protocolName -match '^URL:(.+)$') {
                $protocolName = $Matches[1]
            }

            # Get the URL protocol command from the shell\open\command subkey
            $command = $null
            if ($protocolData.SubKeys.ContainsKey('shell') -and 
                $protocolData.SubKeys['shell'].SubKeys.ContainsKey('open') -and 
                $protocolData.SubKeys['shell'].SubKeys['open'].SubKeys.ContainsKey('command')) {
                $command = $protocolData.SubKeys['shell'].SubKeys['open'].SubKeys['command']['(Default)']
            }

            # Get the URL protocol icon from the DefaultIcon subkey
            $iconPath = $null
            if ($protocolData.SubKeys.ContainsKey('DefaultIcon')) {
                $iconPath = $protocolData.SubKeys['DefaultIcon']['(Default)']
            }

            # Determine if the protocol is built or from Microsoft.
            $isMicrosoft = $protocol -match '^ms-|^microsoft'

            # Create a new object to store the URL protocol data
            $urlProtocol = [PSCustomObject]@{
                Protocol = $protocol
                Name = $protocolName
                Command = $command
                IconPath = $iconPath
                IsMicrosoft = $isMicrosoft
                # Create empty properties for package data later
                PackageName = ""
                PackageAppName = ""
                PackageAppDescription = ""
                ClassID = ""
                PackageAppKeyName = ""
            }

            # Add the URL protocol data object to the array
            $urlProtocolData += $urlProtocol
        }

    #Sort the array by protocol name
    $urlProtocolData = $urlProtocolData | Sort-Object -Property Protocol

    # Gather protocol info from other registry locations to add to the data
    # For installed package Protocols, look in HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages
    # Search app subkeys for Capabilities > URLAssociations. Where entry besides (Default) will be named as the protocol, and will have value of the app class id

    # For debugging
    #$originalProtocolData = Make-DeepCopy $urlProtocolData
    #$regPackageDataOnlyArray = @() # Stores only data from packages, mostly for debugging

    $packagesRegPath = 'Registry::HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages'
    # Get subkeys of the Package key but only where the URLAssociations key exists at Package/WhateverPackage/AppName/Capabilities/URLAssociations and the key has a value besides (Default)
    Get-ChildItem -Path $packagesRegPath -ErrorAction SilentlyContinue | ForEach-Object {
        $packagePath = $_.PSPath
        $packageName = $_.PSChildName
        $packageApps = Get-ChildItem -Path $packagePath
        # Get subkeys of the package
        foreach ($app in $packageApps) {
            $packageAppKeyName = $app.PSChildName
            $capabilitiesPath = Join-Path $app.PSPath "Capabilities"
            $urlAssociationsPath = Join-Path $capabilitiesPath "URLAssociations"


            if (Test-Path $urlAssociationsPath) {
                $urlAssociations = Get-ItemProperty -Path $urlAssociationsPath -ErrorAction SilentlyContinue

                # Get the properties but exclude the built in PSObject properties, leaving only the subkeys that are the protocol names
                $urlAssociations.PSObject.Properties | Where-Object { $_.Name -notin @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider") } | ForEach-Object {
                    $protocol = $_.Name
                    $appClassId = $_.Value

                    # Get the ApplicationName and ApplicationDescription values from Capabilities key value
                    $packageLocalizedAppName = Get-ItemProperty -Path $capabilitiesPath -Name "ApplicationName" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ApplicationName
                    $packageLocalizedAppDescription = Get-ItemProperty -Path $capabilitiesPath -Name "ApplicationDescription" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ApplicationDescription

                    # If the values start with @{ they are a resource and need to be resolved. If they fail just set them blank
                    if ($packageLocalizedAppName -match '^@{') {
                        try{
                            $packageLocalizedAppName = Get-LocalizedString -StringReference $packageLocalizedAppName
                        } catch {
                            $packageLocalizedAppName = ""
                        }
                    }
                    if ($packageLocalizedAppDescription -match '^@{') {
                        try {
                            $packageLocalizedAppDescription = Get-LocalizedString -StringReference $packageLocalizedAppDescription
                        } catch {
                            $packageLocalizedAppDescription = ""
                        }
                    }

                    # Check if the protocol is already in the list, if so merge the data, otherwise create a new object to add to $urlProtocolData
                    if ($urlProtocolData.Protocol -contains $protocol) {
                        # Find the existing object and update the PackageName and ClassID properties
                        $existingProtocol = $urlProtocolData | Where-Object { $_.Protocol -eq $protocol }
                        $existingProtocol.PackageName = $packageName
                        $existingProtocol.PackageAppName = $packageLocalizedAppName
                        $existingProtocol.PackageAppDescription = $packageLocalizedAppDescription
                        $existingProtocol.ClassID = $appClassId
                        $existingProtocol.PackageAppKeyName = $packageAppKeyName

                         # If the package name starts with Microsoft* or Windows then it's a Microsoft package, update the IsMicrosoft property
                        if ($packageName -match '^Microsoft|^Windows') {
                            $existingProtocol.IsMicrosoft = $true
                        }
                    } else {
                        # Determine if microsoft by checking protocol name
                        $isMicrosoft = $false
                        if ($protocol -match '^ms-|^microsoft', 'IgnoreCase') {
                            $isMicrosoft = $true
                        }
                        # If the package name starts with Microsoft* or Windows then it's a Microsoft package
                        if ($packageName -match '^Microsoft|^Windows') {
                            $isMicrosoft = $true
                        }
                        $packageProtocolData = $null
                        # Create a new object to store the URL protocol data
                        $packageProtocolData = [PSCustomObject]@{
                            Protocol = $protocol
                            Name = $packageLocalizedAppName
                            Command = ""
                            IconPath = ""
                            PackageName = $packageName
                            PackageAppKeyName = $packageAppKeyName
                            PackageAppName = $packageLocalizedAppName
                            PackageAppDescription = $packageLocalizedAppDescription
                            ClassID = $appClassId
                            IsMicrosoft = $isMicrosoft
                        }
                        #$regPackageDataOnlyArray += $packageProtocolData # For debugging
                        $urlProtocolData += $packageProtocolData
                    }
                }
            }
        }

    }
    # Now have all data in $urlProtocolData array, but we want to filter it
    # Will only include protocols that are marked as Microsoft or are associated with a package
    $filteredUrlProtocolData = $urlProtocolData | Where-Object { $_.IsMicrosoft -or $_.PackageName }

    return $filteredUrlProtocolData
}

function Create-URL-Protocols-CSVFile {
    param (
        [string]$outputPath,
        [array]$urlProtocolsData
    )

    $csvContent = "Protocol,Name,Command,IconPath,PackageName,PackageAppName,PackageAppDescription,ClassID,IsMicrosoft`n"

    foreach ($item in $urlProtocolsData) {
        $protocol = $item.Protocol -replace '"', '""'
        # Add :// to end of protocol to make it a valid URL
        $protocol += "://"

        $name = $item.Name -replace '"', '""'
        $command = $item.Command -replace '"', '""'
        $iconPath = if ($item.IconPath) {
            "`"$($item.IconPath -replace '"', '""')`""  # Escape double quotes in the icon path.
        } else {
            "None"
        }
        $packageName = $item.PackageName -replace '"', '""'
        $packageAppName = $item.PackageAppName -replace '"', '""'
        $packageAppDescription = $item.PackageAppDescription -replace '"', '""'
        $classID = $item.ClassID -replace '"', '""'
        $isMicrosoft = $item.IsMicrosoft

        $csvContent += "`"$protocol`",`"$name`",`"$command`",$iconPath,`"$packageName`",`"$packageAppName`",`"$packageAppDescription`",`"$classID`",$isMicrosoft`n"
    }

    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Create-Protocol-Shortcut{
    param (
        [string]$protocol,
        [string]$name,
        [string]$command,
        [string]$iconPath,
        [string]$shortcutPath
    )
    Write-Verbose "Creating shortcut for $protocol"
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $protocol + "://"
        #$shortcut.Arguments = $protocol
        if ($iconPath) {
            Write-Verbose "Setting custom icon: $iconPath"
            # If it starts with @{ then it's a resource and needs to be resolved
            if ($iconPath -match '^@{') {
                try {
                    $resolvedIconPath = Get-LocalizedString -StringReference $iconPath
                    Write-Verbose "Resolved icon path: $resolvedIconPath"
                    $shortcut.IconLocation = $resolvedIconPath
                } catch {
                    Write-Verbose "Failed to resolve icon path: $iconPath"
                    Write-Verbose "Using default icon"
                    $shortcut.IconLocation = $iconPath
                }
            }
        } else {
            Write-Verbose "No iconPath provided, using default"
        }

        $shortcut.Save()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $true
    }
    catch {
        Write-Host "Error creating shortcut for $name`: $($_.Exception.Message)"
        return $false
    }
}

# ---------------------------------------------- ----------------------------------------------------------------
# ----------------------------------------------    Main Script    ----------------------------------------------
# ---------------------------------------------- ----------------------------------------------------------------

# Create empty arrays for each type of data to be stored
$clsidInfo = @()
$namedFolders = @()
$taskLinks = @()
$settingsData = @()
$msSettingsList = @()
$deepLinkData = @()
$deepLinksProcessedData = @()
$URLProtocolsData = @()

if (-not $SkipURLProtocols){
    Write-Host "`n----- Processing URL Protocols -----"
    $URLProtocolsData = Get-And-Process-URL-Protocols
    #Write-Host "Found $($URLProtocolsData.Count) URL Protocols"
    foreach ($protocol in $URLProtocolsData) {
        $success = Create-Protocol-Shortcut -protocol $protocol.Protocol -name $protocol.Name -command $protocol.Command -iconPath $protocol.IconPath -shortcutPath (Join-Path $URLProtocolLinksOutputFolder "$($protocol.Protocol).lnk")
        if ($success) {
            Write-Host "Created URL Protocol Shortcut: $($protocol.Name)"
        } else {
            Write-Host "Failed to create URL Protocol shortcut: $($protocol.Name)"
        }
    }
}

if (-not $SkipMSSettings) {
    $msSettingsList = Get-MS-SettingsFrom-SystemSettingsDLL -DllPath $SystemSettingsDllPath

    if ($null -eq $settingsData) {
        Write-Host "No MS Settings data found or error occurred while parsing."
        return
    }

    Write-Host "`n----- Processing MS Settings -----"

    foreach ($setting in $msSettingsList) {
        # Get the short name from the second part of the ms-settings: link
        $settingName = $setting.Split(':')[1]
        # Use policyId as the shortcut name
        $fullShortcutName = $setting

        # Sanitize the shortcut name (although policyId should already be safe)
        $sanitizedName = $settingName -replace '[\\/:*?"<>|]', '_'

        $shortcutPath = Join-Path $msSettingsOutputFolder "$sanitizedName.lnk"

        $success = Create-MSSettings-Shortcut -fullName $fullShortcutName -shortcutPath $shortcutPath

        if ($success) {
            Write-Host "Created MS Settings Shortcut: $fullShortcutName"
        } else {
            Write-Host "Failed to create shortcut: $fullShortcutName"
        }
    }
}

if (-not $SkipDeepLinks -and $allSettingsXmlPath) {
    # Process other settings data
    $deepLinkData = Get-AllSettings-Data -xmlFilePath $allSettingsXmlPath -SaveXML:$SaveXML

    if ($null -eq $deepLinkData) {
        Write-Host "No MS Settings data found or error occurred while parsing."
        return
    }

    Write-Host "`n----- Processing Deep Links -----"

    foreach ($deepLink in $deepLinkData) {
        # Check if it has a DeepLink
        if ($deepLink.DeepLink) {
            $result = Create-Deep-Link-Shortcut -settingArray $deepLink

            if ($result) {
                Write-Host "Created Deep Link Shortcut: $($deepLink.Description)"
                # Add the updated deepLink object to the processed data array. Will also now contain FullCommand and ShortcutPath
                $deepLinksProcessedData += $result
            } else {
                Write-Host "Failed to create Deep Link shortcut: $($deepLink.Description)"
            }
        }
    }
}

# If statement to check if CLSID is skipped by argument
if (-not $SkipCLSID) {
    # Retrieve all CLSIDs with a "ShellFolder" subkey from the registry.
    # These CLSIDs represent shell folders that are embedded within Windows.
    $shellFolders = Get-ChildItem -Path 'Registry::HKEY_CLASSES_ROOT\CLSID' |
    Where-Object {$_.GetSubKeyNames() -contains "ShellFolder"} |
    Select-Object PSChildName

    Write-Host "`n----- Processing $($shellFolders.Count) Shell Folders -----"

    # Loop through each relevant CLSID that was found and process it to create shortcuts.
    foreach ($folder in $shellFolders) {
        $clsid = $folder.PSChildName  # Extract the CLSID.
        Write-Verbose "Processing CLSID: $clsid"

        # Retrieve the name of the shell folder using the Get-FolderName function and the source of the name within the registry
        $resultArray = Get-FolderName -clsid $clsid -CustomLanguageFolder $CustomLanguageFolderPath
        $name = $resultArray[0]
        $nameSource = $resultArray[1]

        # Sanitize the folder name to make it a valid filename.
        $sanitizedName = $name -replace '[\\/:*?"<>|]', '_'
        $shortcutPath = Join-Path $CLSIDshortcutsOutputFolder "$sanitizedName.lnk"

        Write-Verbose "Attempting to create shortcut: $shortcutPath"
        $success = Create-Shortcut -clsid $clsid -name $name -shortcutPath $shortcutPath

        if ($success) {
            Write-Host "Created CLSID Shortcut For: $name"
        }
        else {
            Write-Host "Failed to create shortcut for $name"
        }

        # Check for sub-items (pages) related to the current CLSID (e.g., control panel items).
        $appName = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid" -ErrorAction SilentlyContinue)."System.ApplicationName"

        # Store the CLSID information for later use (e.g., in CSV file generation).
        $iconPath = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\DefaultIcon" -ErrorAction SilentlyContinue).'(default)'
        $clsidInfo += [PSCustomObject]@{
            CLSID = $clsid
            Name = $name
            NameSource = $nameSource
            IconPath = $iconPath
        }
    }

}


if (-not $SkipNamedFolders) {
    # Retrieve all named special folders from the registry.
    $namedFolders = Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions"
    Write-Host "`n----- Processing $($namedFolders.Count) Special Named Folders -----"

    # Loop through each named folder and create a shortcut for it.
    foreach ($folder in $namedFolders) {
        $folderProperties = Get-ItemProperty -Path $folder.PSPath
        $folderName = $folderProperties.Name  # Extract the name of the folder.
        $iconPath = $folderProperties.Icon  # Extract the custom icon path (if any).

        if ($folderName) {
            Write-Verbose "Processing named folder: $folderName"

            # Sanitize the folder name to make it a valid filename.
            $sanitizedName = $folderName -replace '[\\/:*?"<>|]', '_'
            $shortcutPath = Join-Path $namedShortcutsOutputFolder "$sanitizedName.lnk"

            Write-Verbose "Attempting to create shortcut: $shortcutPath"
            $success = Create-NamedShortcut -name $folderName -shortcutPath $shortcutPath -iconPath $iconPath

            if ($success) {
                Write-Host "Created Shortcut For Named Folder: $folderName"
            }
            else {
                Write-Host "Failed to create shortcut for named folder $folderName"
            }
        }
        else {
            Write-Verbose "Skipping folder with no name: $($folder.PSChildName)"
        }
    }
}

if (-not $SkipTaskLinks) {
    # Process Task Links - Use the extracted XML data from Shell32 to create shortcuts for task links
    Write-Host "`n -----Processing Task Links -----"
    # Retrieve task links from the XML content in shell32.dll.
    $taskLinks = Get-TaskLinks -SaveXML:$SaveXML -DLLPath:$CustomDLLPath -CustomLanguageFolder:$CustomLanguageFolderPath
    $createdShortcutNames = @{} # Track created shortcuts to be able to tasks with the same name but different commands by appending a number

    foreach ($task in $taskLinks) {
        $originalName = $task.Name
        $sanitizedName = ""

        # Prepend category names to tasks if DontGroupTasks is not specified
        if (-not $DontGroupTasks) {
            if ($UseAlternativeCategoryNames) {
                # Try looking up the application ID default name by CLSID in the registry
                $trueApplicationName = (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\$($task.ApplicationId)" -ErrorAction SilentlyContinue)."(Default)"
                if ($trueApplicationName) {
                    # If found then prepend it to the task name and use that
                    $sanitizedName = "$trueApplicationName - $originalName"
                }
            }

            # If not set to use alternative categories, or one was not found, then use the default Prettify-App-Name function
            if ($task.ApplicationName -and -not $sanitizedName) {
                # Use Prettify-App-Name function by default, unless DontGroupTasks is specified
                $sanitizedName = Prettify-App-Name -AppName $task.ApplicationName -TaskName $originalName
            } elseif (-not $task.ApplicationName -and -not $sanitizedName) {
                $sanitizedName = $originalName -replace '[\\/:*?"<>|]', '_'
            }
        } else {
            $sanitizedName = $originalName -replace '[\\/:*?"<>|]', '_'
        }

        # Check if a shortcut with this name already exists, if so set a unique number to the end of the name
        $nameCounter = 1
        $uniqueName = $sanitizedName
        while ($createdShortcutNames.ContainsKey($uniqueName)) {
            $nameCounter++
            $uniqueName = "${sanitizedName} ($nameCounter)"
        }

        # Determine the command based on available information. Some task XML entries have the entire command already given, others are implied to be used with control.exe
        if ($task.Command) {
            $command = $task.Command
        } elseif ($task.ApplicationName -and $task.Page) {
            $command = "control.exe /name $($task.ApplicationName) /page $($task.Page)"
        } else {
            Write-Verbose "Skipping task $originalName due to insufficient command information"
            continue
        }


        # Determine whether to create a URL or LNK shortcut based on the command
        $shortcutType = ""
        # If the command starts with a protocol (e.g., http://, https://, or even mshelp://), create a URL shortcut
        if ($command -match '^[a-zA-Z0-9]+:\/\/') {
            $shortcutType = "url"
        # Match ms-protocol shortcuts such as ms-settings: or ms-availablenetworks:, and create a URL if so
        #} elseif ($command -match '^ms-[a-zA-Z0-9]+:') {
        #    $shortcutType = "url"
        # Otherwise default to lnk
        } else {
            $shortcutType = "lnk"
        }

        $shortcutPath = Join-Path $taskLinksOutputFolder "$uniqueName.$shortcutType"
        $createdShortcutNames[$uniqueName] = $true

        $success = Create-TaskLink-Shortcut -name $uniqueName -shortcutPath $shortcutPath -shortcutType $shortcutType -command $command -controlPanelName $task.ControlPanelName -applicationId $task.ApplicationId -keywords $task.Keywords

        if ($success) {
            Write-Host "Created task link shortcut for $uniqueName"
        } else {
            Write-Host "Failed to create task link shortcut for $uniqueName"
        }
    }
}

# Create the CSV files using the stored CLSID and 'task link' (aka menu sub-pages) data. Skip each depending on the corresponding switch.
# Also construct strings containing display paths with nicer formatting to display at the end of the script.
$CLSIDDisplayPath = ""
$namedFolderDisplayPath = ""
$taskLinksDisplayPath = ""
$msSettingsDisplayPath = ""
$deepLinksDisplayPath = ""
$urlProtocolsDisplayPath = ""

if ($SaveCSV) {
    if (-not $SkipCLSID) {
        Create-CLSIDCsvFile -outputPath $clsidCsvPath -clsidData $clsidInfo
        $CLSIDDisplayPath = "`n  > " + $clsidCsvPath
    }

    if (-not $SkipNamedFolders) {
        Create-NamedFoldersCsvFile -outputPath $namedFoldersCsvPath
        $namedFolderDisplayPath = "`n  > " + $namedFoldersCsvPath
    }

    if (-not $SkipTaskLinks) {
        Create-TaskLinksCsvFile -outputPath $taskLinksCsvPath -taskLinksData $taskLinks
        $taskLinksDisplayPath = "`n  > " + $taskLinksCsvPath
    }
    if (-not $SkipMSSettings) {
        Create-MSSettingsCsvFile -outputPath $msSettingsCsvPath -msSettingsList $msSettingsList
        $msSettingsDisplayPath = "`n  > " + $msSettingsCsvPath
    }
    if (-not $SkipDeepLinks) {
        Create-Deep-Link-CSVFile -outputPath $deepLinksCsvPath -deepLinksDataArray $deepLinksProcessedData
        $deepLinksDisplayPath = "`n  > " + $deepLinksCsvPath
    }
    if (-not $SkipURLProtocols) {
        Create-URL-Protocols-CSVFile -outputPath $URLProtocolLinksCsvPath -urlProtocolsData $URLProtocolsData
        $urlProtocolsDisplayPath = "`n  > " + $URLProtocolLinksCsvPath
    }
}


# Output a message indicating that the script execution is complete and the CSV files have been created.
Write-Host "`n-----------------------------------------------"
Write-Host "        Super God Mode Script Complete" -ForeGroundColor Yellow
Write-Host "-----------------------------------------------`n"

# Display total counts
$totalCount = $clsidInfo.Count + $namedFolders.Count + $taskLinks.Count + $msSettingsList.Count + $deepLinksProcessedData.Count + $URLProtocolsData.Count

# Output the total counts of each, and color the numbers to stand out. Done by writing the text and then the number separately with -NoNewLine. If it was skipped, also add that but not colored.
Write-Host "Total Shortcuts Created: " -NoNewline
Write-Host $totalCount -ForegroundColor Green

Write-Host "  > CLSID Links:     " -NoNewline
Write-Host $clsidInfo.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipCLSID) { "   (Skipped)" }) # If skipped, add the skipped text, otherwise still write empty string because we used -NoNewline previously

Write-Host "  > Named Folders:   " -NoNewline
Write-Host $namedFolders.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipNamedFolders) { "   (Skipped)" })

Write-Host "  > Task Links:      " -NoNewline
Write-Host $taskLinks.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipTaskLinks) { "   (Skipped)" })

Write-Host "  > Settings Links:  " -NoNewline
Write-Host $msSettingsList.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipMSSettings) { "   (Skipped)" })

Write-Host "  > Deep Links:      " -NoNewline
Write-Host $deepLinksProcessedData.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipDeepLinks) { "   (Skipped)" })

Write-Host "  > URL Protocols:   " -NoNewline
Write-Host $URLProtocolsData.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipURLProtocols) { "   (Skipped)" })

Write-Host "`n-----------------------------------------------`n"

# If SaveXML switch was used, also output the paths to the saved XML files
if ($SaveXML -and -not $SkipTaskLinks) {
    Write-Host "XML & txt files have been saved at:"
    if (-not $SkipTaskLinks) {
        Write-Host "  > $xmlContentFilePath"
        Write-Host "  > $resolvedXmlContentFilePath"
    }
    if (-not $SkipDeepLinks) {
        Write-Host "  > $resolvedSettingsXmlContentFilePath"
    }
    Write-Host "`n"
}

# As long as any of the CSV files were created, output the paths to them - check by seeing if strings are empty
if ($CLSIDDisplayPath -or $namedFolderDisplayPath -or $taskLinksDisplayPath -or $msSettingsDisplayPath -or $deepLinksDisplayPath) {
    $csvPrintString = "CSV files have been created at:" + "$CLSIDDisplayPath" + "$namedFolderDisplayPath" + "$taskLinksDisplayPath" + "$msSettingsDisplayPath" + "$deepLinksDisplayPath" + "$urlProtocolsDisplayPath" + "`n"
    Write-Host $csvPrintString
}
