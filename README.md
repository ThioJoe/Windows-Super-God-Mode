# Windows "Super God Mode"

This PowerShell script creates shortcuts to all special shell folders, named folders, and task links in Windows, providing easy access to a wide range of system settings and features. It was inspired by the famously nicknamed "God Mode" folder, and creates many more shortcuts than even that.

## Features

- Creates shortcuts for all CLSID-based shell folders
- Creates shortcuts for all named special folders
- Creates shortcuts for task links (sub-pages within shell folders and control panel menus)
- Generates CSV files with detailed information about the shortcuts
- Optionally saves XML content from shell32.dll for reference
- Customizable output location and various execution options

## Usage

1. Download the latest version of the script. (Direct link [here](https://github.com/ThioJoe/Windows-Super-God-Mode/raw/main/Super_God_Mode.ps1))
2. Open PowerShell to the directory with the script. (Tip: In File Explorer, just type "PowerShell.exe" into the address bar to open it to that path).
3. Run the following command to allow script execution for the current session:
   ```
   Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process
   ```
5. Run the script with any desired optional parameters:
   ```
   .\Super_God_Mode.ps1 [optional parameters]
   ```

### Parameters

- `-SaveCSV`: Save information about all shortcuts into CSV files
- `-SaveXML`: Save the XML content from shell32.dll
- `-Output`: Specify a custom output folder path
- `-DeletePreviousOutputFolder`: Delete the previous output folder before running
- `-Verbose`: Enable verbose output
- `-DontGroupTasks`: Prevent grouping task shortcuts by application name
- `-SkipCLSID`: Skip creating shortcuts for CLSID-based shell folders
- `-SkipNamedFolders`: Skip creating shortcuts for named special folders
- `-SkipTaskLinks`: Skip creating shortcuts for task links
- `-DLLPath`: Specify a custom DLL file path for shell32.dll

### Example

```powershell
.\Super_God_Mode.ps1 -SaveCSV -SaveXML -Output "C:\SuperGodMode" -Verbose
```

## Output

The script creates a folder (default name: "Shell Folder Shortcuts") containing:

- Shortcuts to CLSID-based shell folders
- Shortcuts to named special folders
- Shortcuts to task links
- CSV files with detailed information about the shortcuts (if `-SaveCSV` is used)
- XML files with shell32.dll content (if `-SaveXML` is used)

## Notes

- Some shortcuts may not work on all Windows versions due to differences in available features.
- The script does not modify any system settings; it only creates shortcuts to existing Windows features.

___

# Extra Tools

The "Extra Tools" folder contains additional scripts that complement the main functionality of Windows Super God Mode:

### Get_DLL_String_Reference.ps1

This script allows you to easily retrieve the localized string of a single specific string reference.

**Features:**
- Interactively prompts for string references
- Resolves and displays the localized string values
- Supports the `@dllpath,-resourceID` format

**Usage:**
1. Run the script in PowerShell
2. Enter the string reference when prompted (e.g., `@%SystemRoot%\system32\shell32.dll,-9227`)
3. The script will display the resolved string value

### Windows_XML_String_Resolver.ps1

This script processes entire XML files containing Windows string references and resolves them to their actual string values. Mostly intended to be used with the XML from shell32.dll.mun containing all the Windows task links.

**Features:**
- Processes entire XML files, replacing string references with their resolved values
- Supports custom DLL paths for resolving strings
- Generates a new XML file with resolved strings

**Usage:**
```powershell
.\Windows_XML_String_Resolver.ps1 -XmlFilePath "path\to\your\file.xml" [-CustomResourcePaths "shell32=C:\custom\path\shell32.dll", "user32=C:\another\path\user32.mui"] [-Debug]
```
