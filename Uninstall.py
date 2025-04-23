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
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    shortcut_path = os.path.join(desktop_path, "Workplace Scheduler.bat")
    
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
        print(f"Removed: {shortcut_path}")
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
            shutil.rmtree(dir_path)
            print(f"Removed directory: {dir_path}")
    
    # Remove data file
    data_file = os.path.join(app_dir, "data.json")
    if os.path.exists(data_file):
        os.remove(data_file)
        print(f"Removed data file: {data_file}")
    
    print_header("UNINSTALLATION COMPLETE")
    print("\nThe application has been uninstalled. You may safely delete")
    print("the remaining application files manually if desired.")
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    main()
