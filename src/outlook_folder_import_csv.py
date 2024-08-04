import json
import sys
from pathlib import Path
from typing import Any, Dict

import win32com.client


def create_folder_structure(
    parent_folder: Any, folder_data: Dict[str, Any]
) -> None:
    folder_name = folder_data["name"]
    try:
        # Check if the folder already exists
        existing_folder = parent_folder.Folders(folder_name)
        print(f"Folder '{folder_name}' already exists. Skipping.")
    except:
        # Create a new folder if it doesn't exist
        existing_folder = parent_folder.Folders.Add(folder_name)
        print(f"Created folder '{folder_name}'.")

    # Recursively create subfolders
    for subfolder_data in folder_data["subfolders"]:
        create_folder_structure(existing_folder, subfolder_data)


def import_folders(account_name: str, json_file: str) -> None:
    json_path = Path(json_file).resolve()
    if not json_path.exists():
        print(f"File '{json_file}' not found.")
        return

    try:
        with json_path.open("r", encoding="utf-8") as f:
            folder_structure: Dict[str, Any] = json.load(f)
    except json.JSONDecodeError:
        print(f"Error: '{json_file}' is not a valid JSON file.")
        return
    except Exception as e:
        print(f"Error reading file '{json_file}': {e}")
        return

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
        "MAPI"
    )

    # Search for the account
    root_folder = None
    for account in outlook.Accounts:
        if account.DisplayName == account_name:
            root_folder = account.DeliveryStore.GetRootFolder()
            break

    if root_folder is None:
        print(f"Account '{account_name}' not found.")
        return

    create_folder_structure(root_folder, folder_structure)
    print("Folder structure import completed.")


if __name__ == "__main__":
    args_len = 3
    if len(sys.argv) != args_len:
        print(
            "Usage: python import_script.py <Outlook account name> <JSON file name>"
        )
    else:
        import_folders(sys.argv[1], sys.argv[2])
