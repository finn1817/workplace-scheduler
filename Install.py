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

def main():
    print_header("WORKPLACE SCHEDULER INSTALLER")
    
    # Check Python version
    print_step("Checking Python version...")
    if sys.version_info < (3, 7):
        print("Error: Python 3.7 or higher is required.")
        sys.exit(1)
    print(f"Python version {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} detected.")
    
    # Install required packages
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
    
    # Create application directories
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
    
    # Create initial data file
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
    
    # Create desktop shortcut
    print_step("Creating desktop shortcut...")
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    shortcut_path = os.path.join(desktop_path, "Workplace Scheduler.bat")
    
    with open(shortcut_path, 'w') as f:
        f.write(f'@echo off\n"{sys.executable}" "{os.path.join(app_dir, "App.py")}"\npause')
    
    print(f"Created desktop shortcut: {shortcut_path}")
    
    print_header("INSTALLATION COMPLETE")
    print("\nYou can now run the application by double-clicking the")
    print("'Workplace Scheduler' shortcut on your desktop.")
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    main()
