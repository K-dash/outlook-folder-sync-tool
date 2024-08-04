# Outlook Folder Sync Tool

This tool allows you to export and import Outlook folder structures, facilitating easy synchronization or migration of folder organizations between different Outlook accounts.

## Features

- Export Outlook folder structure to a JSON file
- Import Outlook folder structure from a JSON file
- Supports nested folder hierarchies
- Command-line interface for easy integration into workflows

## Requirements

- Windows operating system
- Python 3.7 or higher
- Poetry for dependency management
- Microsoft Outlook installed on the system

**Note**: This tool is designed to work exclusively on Windows environments due to its dependency on the Win32 COM interface for Outlook.

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/K-dash/outlook-folder-sync-tool.git
   cd outlook-folder-sync-tool
   ```

2. Install dependencies using Poetry:
   ```
   poetry install
   ```

## Usage

### Using Python (with Poetry)

1. To export a folder structure:
   ```
   poetry run python src/outlook_folder_export_csv.py <Outlook account name> <Folder name to export>
   ```

2. To import a folder structure:
   ```
   poetry run python src/outlook_folder_import_csv.py <Outlook account name> <JSON file name>
   ```

### Using Executable Files

1. Create executable files using PyInstaller:
   ```
   poetry run pyinstaller --onefile src/outlook_folder_export_csv.py
   poetry run pyinstaller --onefile src/outlook_folder_import_csv.py
   ```

   Alternatively, you can create both executables at once using:
   ```
   poetry run pyinstaller --onefile src/outlook_folder_export_csv.py src/outlook_folder_import_csv.py
   ```

2. The executable files will be created in the `dist` directory.

3. To export a folder structure:
   ```
   dist\outlook_folder_export_csv.exe <Outlook account name> <Folder name to export>
   ```

4. To import a folder structure:
   ```
   dist\outlook_folder_import_csv.exe <Outlook account name> <JSON file name>
   ```

## Examples

1. Exporting the folder structure of the "Inbox" folder for the account "john.doe@example.com":
   ```
   poetry run python src/outlook_folder_export_csv.py john.doe@example.com Inbox
   ```

2. Importing a folder structure from "folder_structure.json" to the account "jane.doe@example.com":
   ```
   poetry run python src/outlook_folder_import_csv.py jane.doe@example.com folder_structure.json
   ```

3. Example of a JSON file structure (folder_structure.json):
   ```json
   {
     "name": "Inbox",
     "subfolders": [
       {
         "name": "Project A",
         "subfolders": [
           {
             "name": "Meetings",
             "subfolders": []
           },
           {
             "name": "Documents",
             "subfolders": []
           }
         ]
       },
       {
         "name": "Personal",
         "subfolders": []
       },
       {
         "name": "Archive",
         "subfolders": [
           {
             "name": "2022",
             "subfolders": []
           },
           {
             "name": "2023",
             "subfolders": []
           }
         ]
       }
     ]
   }
   ```

   This JSON structure represents an "Inbox" with subfolders for "Project A" (which has its own subfolders), "Personal", and an "Archive" folder with yearly subfolders.

## Troubleshooting

- Ensure that Outlook is installed and properly configured on your system.
- If you encounter permission issues, try running the script or executable as an administrator.
- Make sure the specified Outlook account exists and is accessible.

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## WIP
- Add unit test.
- Export/Import account-independent mail distribution rules.
- `.exe` file package distribution.
- Add pre-commit hook.
- Add CI pipline
