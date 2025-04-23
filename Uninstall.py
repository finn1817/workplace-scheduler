import os
import sys
import shutil
from pathlib import Path

def print_header(text):
    print("\n" + "=" * 60)
    print(f" {text} ".center(60, "="))
    print("=" * 60)

def print_step(text):
    print(f"\n>> {text}")

def find_desktop_path():
    """Find the correct desktop path, handling OneDrive scenarios"""
    # Standard desktop path
    standard_path = os.path.join(os.path.expanduser("~"), "Desktop")
    
    # OneDrive desktop path
    onedrive_path = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
    
    # Check which one exists
    if os.path.exists(onedrive_path):
        return onedrive_path
    elif os.path.exists(standard_path):
        return standard_path
    else:
        return standard_path  # Return standard path even if it doesn't exist

def main():
    print_header("WORKPLACE SCHEDULER UNINSTALLER")
    
    # Get confirmation
    print("\nWARNING: This will remove all application data including:")
    print("  - All workplace data")
    print("  - All saved schedules")
    print("  - All configuration settings")
    
    confirm = input("\nAre you sure you want to continue? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("\nUninstallation cancelled.")
        input("\nPress Enter to exit...")
        return
    
    # Remove desktop shortcut
    print_step("Removing desktop shortcut...")
    desktop_path = find_desktop_path()
    shortcut_path = os.path.join(desktop_path, "Workplace Scheduler.bat")
    
    if os.path.exists(shortcut_path):
        try:
            os.remove(shortcut_path)
            print(f"Removed: {shortcut_path}")
        except Exception as e:
            print(f"Warning: Could not remove desktop shortcut: {str(e)}")
    else:
        print("Desktop shortcut not found.")
    
    # Remove data directories
    print_step("Removing application data...")
    app_dir = os.path.dirname(os.path.abspath(__file__))
    
    directories = [
        "workplaces",
        "schedules",
        "saved_schedules",
        "templates",
        "static",
        "logs"
    ]
    
    for directory in directories:
        dir_path = os.path.join(app_dir, directory)
        if os.path.exists(dir_path):
            try:
                shutil.rmtree(dir_path)
                print(f"Removed directory: {dir_path}")
            except Exception as e:
                print(f"Warning: Could not remove directory {dir_path}: {str(e)}")
    
    # Remove data file
    data_file = os.path.join(app_dir, "data.json")
    if os.path.exists(data_file):
        try:
            os.remove(data_file)
            print(f"Removed data file: {data_file}")
        except Exception as e:
            print(f"Warning: Could not remove data file: {str(e)}")
    
    print_header("UNINSTALLATION COMPLETE")
    print("\nThe application has been uninstalled. You may safely delete")
    print("the remaining application files manually if desired.")
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    main()
