"""
Check Folder Access Script
ตรวจสอบว่าระบบสามารถเข้าถึง folder ที่กำหนดได้หรือไม่
"""

import os
import sys
import platform


def check_folder_access(folder_path: str) -> dict:
    """ตรวจสอบสิทธิ์การเข้าถึง folder"""
    result = {
        "path": folder_path,
        "exists": False,
        "readable": False,
        "writable": False,
        "is_directory": False,
        "files_count": None,
        "error": None,
    }

    try:
        result["exists"] = os.path.exists(folder_path)

        if not result["exists"]:
            result["error"] = f"Path does not exist: {folder_path}"
            return result

        result["is_directory"] = os.path.isdir(folder_path)

        if not result["is_directory"]:
            result["error"] = f"Path is not a directory: {folder_path}"
            return result

        result["readable"] = os.access(folder_path, os.R_OK)
        result["writable"] = os.access(folder_path, os.W_OK)

        if result["readable"]:
            files = os.listdir(folder_path)
            result["files_count"] = len(files)

    except PermissionError as e:
        result["error"] = f"Permission denied: {e}"
    except OSError as e:
        result["error"] = f"OS error: {e}"

    return result


def print_result(result: dict):
    """แสดงผลลัพธ์การตรวจสอบ"""
    print("=" * 60)
    print("  Folder Access Check Result")
    print("=" * 60)
    print(f"  Path:         {result['path']}")
    print(f"  Exists:       {'Yes' if result['exists'] else 'No'}")
    print(f"  Is Directory: {'Yes' if result['is_directory'] else 'No'}")
    print(f"  Readable:     {'Yes' if result['readable'] else 'No'}")
    print(f"  Writable:     {'Yes' if result['writable'] else 'No'}")

    if result["files_count"] is not None:
        print(f"  Files/Folders: {result['files_count']} items")

    if result["error"]:
        print(f"  Error:        {result['error']}")

    print("=" * 60)

    if result["readable"] and result["is_directory"]:
        print("  Status: ACCESS OK")
    else:
        print("  Status: ACCESS FAILED")

    print("=" * 60)


def main():
    # Default target folder
    target_folder = r"N:\1.ThailandTrophy\7.Catalog"

    # Allow overriding via command line argument
    if len(sys.argv) > 1:
        target_folder = sys.argv[1]

    print(f"\nSystem: {platform.system()} {platform.release()}")
    print(f"Checking access to: {target_folder}\n")

    result = check_folder_access(target_folder)
    print_result(result)

    # Return exit code: 0 = accessible, 1 = not accessible
    sys.exit(0 if result["readable"] and result["is_directory"] else 1)


if __name__ == "__main__":
    main()
