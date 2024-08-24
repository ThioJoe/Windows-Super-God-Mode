# "Super God Mode" For Windows

This PowerShell script <b>creates shortcuts to all special shell folders, named folders, task links, system settings, deep links, and URL protocols in Windows</b>, providing easy access to a wide range of system settings and features.

It was inspired by the famously nicknamed "God Mode" folder and creates many more shortcuts than even that.

## Screenshots


<p align="center">
<img width="700" alt="GUI Window" src="https://github.com/user-attachments/assets/2103b265-d2e5-4fa7-ac69-362784bcb0db">
</p><p align="center">
<img width="290" alt="Results" src="https://github.com/user-attachments/assets/4d01fbad-b597-4433-bd67-2638ded8a6ed">
<img width="392" alt="Output Folders" src="https://github.com/user-attachments/assets/898efc48-ddc6-4875-b906-b89963d5778e">
</p>



## Features

- Creates shortcuts for various Windows components:
  - CLSID-based shell folders
  - Named special folders
  - Task links (sub-pages within shell folders and control panel menus)
  - System settings (ms-settings: links)
  - Deep links (direct links to various settings menus across Windows)
  - URL protocols
  - Hidden App Links (Internal-use and undocumented URL links used by apps)
- Generates CSV files with detailed information about the shortcuts
- Saves XML content retrieved from shell32.dll and other sources for reference
- Graphical User Interface (GUI) for easy configuration

## How to Run:

### Option 1 (Easier): Using .Bat Launcher
1. Download the latest version of the script. (Direct link [here](https://github.com/ThioJoe/Windows-Super-God-Mode/releases/download/v1.1.0/Super_God_Mode.ps1))
2. Download the launcher batch file to the same location. (Direct link [here](https://github.com/ThioJoe/Windows-Super-God-Mode/releases/download/v1.1.0/SuperGodMode-EasyLauncher.bat))
3. Run the batch file.

### Option 2: Manually running

1. Download the latest version of the script. (Direct link [here](https://github.com/ThioJoe/Windows-Super-God-Mode/releases/download/v1.1.0/Super_God_Mode.ps1))
2. Open PowerShell to the directory with the script. (Tip: In File Explorer, just type "PowerShell.exe" into the address bar to open it to that path).
3. Run the following command to allow script execution for the current session:
   ```
   Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process
   ```
4. Run the script:
   ```
   .\Super_God_Mode.ps1
   ```
   - If no parameters are provided, a GUI will appear for easy configuration.
   - You can also run the script with optional parameters (see below).

## CLI Parameters

Note: Except for `-Debug` and `-Verbose`, you must use `-NoGUI` for arguments to take effect.

#### Alternative Options Arguments

- `-DontGroupTasks`: Prevent grouping task shortcuts by application name
- `-UseAlternativeCategoryNames`: Use alternative category names for task links
- `-AllURLProtocols`: Include third-party URL protocols from installed software
- `-CollectExtraURLProtocolInfo`: Collect additional information about URL protocols

#### Control Output

- `-Output`: Specify a custom output folder path
- `-KeepPreviousOutputFolders`: Don't auto-delete existing output folders before running

#### Arguments to Limit Shortcut Creation

- `-NoStatistics`: Don't create statistics folder and files
- `-NoReadMe`: Don't create tips text file
- `-SkipCLSID`: Skip creating shortcuts for CLSID-based shell folders
- `-SkipNamedFolders`: Skip creating shortcuts for named special folders
- `-SkipTaskLinks`: Skip creating shortcuts for task links
- `-SkipMSSettings`: Skip creating shortcuts for ms-settings: links
- `-SkipDeepLinks`: Skip creating shortcuts for deep links
- `-SkipURLProtocols`: Skip creating shortcuts for URL protocols
- `-SkipHiddenAppLinks`: Skip creating shortcuts to hidden app links

#### Debugging

- `-Verbose`: Enable verbose output. Can be used with or without `-NoGUI`.
- `-Debug`: Enable debug output (also enables verbose output). Can be used with or without `-NoGUI`.

#### Advanced Arguments

- `-NoGUI`: Skip the GUI dialog and run with default or provided parameters
- `-CustomDLLPath`: Specify a custom DLL file path for shell32.dll
- `-CustomLanguageFolderPath`: Specify a path to a folder containing language-specific MUI files
- `-CustomSystemSettingsDLLPath`: Specify a custom path to the SystemSettings.dll file
- `-CustomAllSystemSettingsXMLPath`: Specify a custom path to the "AllSystemSettings_" XML file

### Example

```powershell
.\Super_God_Mode.ps1 -Output "C:\SuperGodMode" -AllURLProtocols -Verbose
```

## Output

The script creates a folder (default name: "Super God Mode") containing:

- Shortcuts to CLSID-based shell folders
- Shortcuts to named special folders
- Shortcuts to task links
- Shortcuts to system settings (ms-settings: links)
- Shortcuts to deep links
- Shortcuts to URL protocols
- Shortcuts to Hidden App Links
- A Statistics folder (With CSV and XML data files)
- A text file with some tips and other info

## Notes

- Some shortcuts may not work on all Windows versions due to differences in available features.
- The script does not modify any system settings; it only creates shortcuts to existing Windows features.
- All parameters and GUI settings are optional. The script will run with default settings if the user doesn't change anything.

___

# Extra Tools

The "Extra Tools" folder contains additional scripts that complement the main functionality of Windows Super God Mode:

### Get_DLL_String_Reference.ps1

This script allows you to easily retrieve the localized string of a single specific string reference.

Features:
- Interactively prompts for string references
- Resolves and displays the localized string values
- Supports the `@dllpath,-resourceID` format

Usage:
1. Run the script in PowerShell
2. Enter the string reference when prompted (e.g., `@%SystemRoot%\system32\shell32.dll,-9227`)
3. The script will display the resolved string value

### Windows_XML_String_Resolver.ps1

This script processes entire XML files containing Windows string references and resolves them to their actual string values. Mostly intended to be used with the XML from shell32.dll.mun containing all the Windows task links.

Features:
- Processes entire XML files, replacing string references with their resolved values
- Supports custom DLL paths for resolving strings
- Generates a new XML file with resolved strings

Usage:
```powershell
.\Windows_XML_String_Resolver.ps1 -XmlFilePath "path\to\your\file.xml" [-CustomResourcePaths "shell32=C:\custom\path\shell32.dll", "user32=C:\another\path\user32.mui"] [-Debug]
```

### Get-MS-Settings-Strings.ps1

This script will find text strings of "ms-settings:" in a DLL file and output them to a text file. 
It is a standalone version of the feature built into the main script. Intended mainly for: "C:\Windows\ImmersiveControlPanel\SystemSettings.dll".

Usage:
```
`.\Get-MS-Settings-Strings.ps1 -DllPath "C:\Windows\ImmersiveControlPanel\SystemSettings.dll" -OutputFilePath "SystemSettings-MS-Settings.txt"
```
- If not specified via arguments, the script will prompt the user for the DLL path, and output to the same directory as the script.

### Find_URLs_From_AppxPackage_Files.ps1

This script fetches the URI protocols for each installed AppxPackage via their AppxManifest.xml file, then brute force searches for those URIs in all files in the app's install directory.
It is a standalone version of the feature built into the main script.

Usage:
- No arguments necessary:  `.\Find_URLs_From_AppxPackage_Files.ps1`
