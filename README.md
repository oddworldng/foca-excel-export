# FocaExcelExport

FOCA plugin to export project metadata (files, URLs, users, locations, emails, client) to Excel.

## Overview

FocaExcelExport is a plugin for FOCA (Forensic Case Analyzer) that allows users to export project metadata to Excel format. The plugin integrates with FOCA's menu system and provides a WinForms dialog to select projects and export their metadata.

## Features

- Export FOCA project metadata to Excel (.xlsx) files
- Dynamic database schema detection
- Progress tracking during export
- User-friendly interface with project selection dropdown

## Requirements

- FOCA (Forensic Case Analyzer)
- .NET Framework 4.7.1 or higher
- Microsoft SQL Server database (used by FOCA)

## Installation

1. Build the FocaExcelExport project
2. Copy the generated `FocaExcelExport.dll` file to FOCA's `Plugins` folder
3. Ensure the ClosedXML dependency is available (either install via NuGet in FOCA's environment or include the DLLs with the plugin)
4. Restart FOCA

## Usage

1. Open FOCA
2. Navigate to the **Export to Excel** option in the menu
3. Select a project from the dropdown list
4. Click **Export** button
5. Choose a destination file using the save dialog
6. The Excel file will be generated with the following columns:
   - Fichero (real file name)
   - URL (URL where the file was found)
   - Usuario (extracted username)
   - Ubicación (network path or file location)
   - Email (email of the user)
   - Cliente (name of the client or machine)

## Building from Source

1. Open the solution in Visual Studio
2. Ensure .NET Framework 4.7.1 is targeted
3. Restore NuGet packages (ClosedXML and DocumentFormat.OpenXml)
4. Build the solution
5. The plugin DLL will be available in the bin directory

## Project Structure

```
foca-excel-export/
├── Classes/
│   ├── ConnectionResolver.cs     # Reads FOCA's database connection string
│   ├── Exporter.cs              # Performs database queries and Excel generation
│   └── SchemaResolver.cs        # Discovers table and column names dynamically
├── Forms/
│   ├── ExportDialog.cs          # Main form with UI controls
│   └── ExportDialog.Designer.cs # Form designer code
├── Properties/
│   └── AssemblyInfo.cs          # Assembly metadata
├── FocaExcelExport.csproj       # Project file
├── FocaExcelExport.sln          # Solution file
├── Plugin.cs                    # Main plugin class implementing FOCA interface
├── packages.config              # NuGet package dependencies
├── README.md                    # This file
└── LICENSE                      # MIT License
```

## License

MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.