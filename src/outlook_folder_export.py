import json
import sys
from typing import Any, Dict, Optional

import win32com.client


def find_folder(root_folder: Any, folder_name: str) -> Optional[Any]:
    if root_folder.Name == folder_name:
        return root_folder
    for folder in root_folder.Folders:
        found = find_folder(folder, folder_name)
        if found:
            return found
    return None


def get_folder_structure(folder: Any) -> Dict[str, Any]:
    structure = {"name": folder.Name, "subfolders": []}
    for subfolder in folder.Folders:
        structure["subfolders"].append(get_folder_structure(subfolder))
    return structure


def export_folders(account_name: str, target_folder_name: str) -> None:
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

    # Search for the specified folder
    target_folder = find_folder(root_folder, target_folder_name)
    if not target_folder:
        print(f"Folder '{target_folder_name}' not found.")
        return

    folder_structure = get_folder_structure(target_folder)

    with open("folder_structure.json", "w", encoding="utf-8") as f:
        json.dump(folder_structure, f, ensure_ascii=False, indent=2)

    print(f"Folder structure of '{target_folder_name}' has been exported.")


if __name__ == "__main__":
    args_len = 3
    if len(sys.argv) != args_len:
        print(
            "Usage: python export_script.py <Outlook account name> <Folder name to export>"
        )
    else:
        export_folders(sys.argv[1], sys.argv[2])
