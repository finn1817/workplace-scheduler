import os
import sys
import subprocess
import shutil
import json
from pathlib import Path

def print_header(text):
    print("\n" + "=" * 60)
    print(f" {text} ".center(60, "="))
    print("=" * 60)

def print_step(text):
    print(f"\n>> {text}")

def create_directory(path):
    if not os.path.exists(path):
        os.makedirs(path)
        print(f"Created directory: {path}")
    else:
        print(f"Directory already exists: {path}")

def find_desktop_path():
    """Find the correct desktop path, handling OneDrive scenarios"""
    # standard desktop path
    standard_path = os.path.join(os.path.expanduser("~"), "Desktop")
    
    # onedrive desktop path
    onedrive_path = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
    
    # check which one exists
    if os.path.exists(onedrive_path):
        return onedrive_path
    elif os.path.exists(standard_path):
        return standard_path
    else:
        # if neither exists, create and use the standard path
        os.makedirs(standard_path, exist_ok=True)
        return standard_path

def main():
    print_header("WORKPLACE SCHEDULER INSTALLER")
    
    # check python version
    print_step("Checking Python version...")
    if sys.version_info < (3, 7):
        print("Error: Python 3.7 or higher is required.")
        sys.exit(1)
    print(f"Python version {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} detected.")
    
    # install required packages
    print_step("Installing required packages...")
    packages = [
        "pandas", 
        "openpyxl", 
        "matplotlib", 
        "PyQt5", 
        "email-validator", 
        "Pillow"
    ]
    
    for package in packages:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    
    # create application directories
    print_step("Creating application directories...")
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
        create_directory(os.path.join(app_dir, directory))
    
    # create initial data file
    print_step("Creating initial data file...")
    data_file = os.path.join(app_dir, "data.json")
    if not os.path.exists(data_file):
        initial_data = {
            "esports_lounge": {
                "hours_of_operation": {
                    "Sunday": [],
                    "Monday": [{"start": "10:00", "end": "22:00"}],
                    "Tuesday": [{"start": "10:00", "end": "22:00"}],
                    "Wednesday": [{"start": "10:00", "end": "22:00"}],
                    "Thursday": [{"start": "10:00", "end": "22:00"}],
                    "Friday": [{"start": "10:00", "end": "22:00"}],
                    "Saturday": []
                }
            },
            "esports_arena": {
                "hours_of_operation": {
                    "Sunday": [],
                    "Monday": [{"start": "10:00", "end": "22:00"}],
                    "Tuesday": [{"start": "10:00", "end": "22:00"}],
                    "Wednesday": [{"start": "10:00", "end": "22:00"}],
                    "Thursday": [{"start": "10:00", "end": "22:00"}],
                    "Friday": [{"start": "10:00", "end": "22:00"}],
                    "Saturday": []
                }
            },
            "it_service_center": {
                "hours_of_operation": {
                    "Sunday": [],
                    "Monday": [{"start": "08:00", "end": "17:00"}],
                    "Tuesday": [{"start": "08:00", "end": "17:00"}],
                    "Wednesday": [{"start": "08:00", "end": "17:00"}],
                    "Thursday": [{"start": "08:00", "end": "17:00"}],
                    "Friday": [{"start": "08:00", "end": "17:00"}],
                    "Saturday": []
                }
            }
        }
        
        with open(data_file, 'w') as f:
            json.dump(initial_data, f, indent=4)
        print(f"Created initial data file: {data_file}")
    else:
        print(f"Data file already exists: {data_file}")
    
    # create desktop shortcut
    print_step("Creating desktop shortcut...")
    desktop_path = find_desktop_path()
    shortcut_path = os.path.join(desktop_path, "Workplace Scheduler.bat")
    
    try:
        with open(shortcut_path, 'w') as f:
            f.write(f'@echo off\ncd /d "{app_dir}"\n"{sys.executable}" "{os.path.join(app_dir, "App.py")}"\npause')
        
        print(f"Created desktop shortcut: {shortcut_path}")
    except Exception as e:
        print(f"Warning: Could not create desktop shortcut: {str(e)}")
        print(f"You can manually create a shortcut to: {os.path.join(app_dir, 'App.py')}")
    
    print_header("INSTALLATION COMPLETE")
    print("\nYou can now run the application by double-clicking the")
    print("'Workplace Scheduler' shortcut on your desktop.")
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    main()
