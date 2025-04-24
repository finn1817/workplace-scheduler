import os
import sys
import json
import re
import random
import logging
import smtplib
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # use non-interactive backend
from datetime import datetime, time, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QPushButton, QComboBox, QTableWidget, QTableWidgetItem,
                            QFileDialog, QMessageBox, QTabWidget, QLineEdit, QCheckBox,
                            QTimeEdit, QSpinBox, QFormLayout, QGroupBox, QTextEdit, QDialog,
                            QScrollArea, QFrame, QSplitter, QStackedWidget, QListWidget,
                            QGridLayout, QHeaderView, QListWidgetItem, QDateEdit, QCalendarWidget)
from PyQt5.QtCore import Qt, QTime, QSize, QSettings, pyqtSignal, QThread, QDate
from PyQt5.QtGui import QFont, QIcon, QPixmap, QColor, QPalette
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog

# constants
DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
APP_NAME = "Workplace Scheduler"
APP_VERSION = "1.0.0"

# get the application directory
APP_DIR = os.path.dirname(os.path.abspath(__file__))

# define directories with absolute paths
DIRS = {
    'workplaces': os.path.join(APP_DIR, 'workplaces'),
    'schedules': os.path.join(APP_DIR, 'schedules'),
    'saved_schedules': os.path.join(APP_DIR, 'saved_schedules'),
    'logs': os.path.join(APP_DIR, 'logs'),
}

# ensure directories exist
for directory in DIRS.values():
    os.makedirs(directory, exist_ok=True)

# data file with absolute path
DATA_FILE = os.path.join(APP_DIR, 'data.json')

# setup logging with absolute path
LOG_FILE = os.path.join(DIRS['logs'], 'app.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# utility functions
def load_data():
    """Load application data from JSON file"""
    if not os.path.exists(DATA_FILE):
        return {}
    try:
        with open(DATA_FILE, 'r') as f:
            return json.load(f)
    except Exception as e:
        logging.error(f"Error loading data: {str(e)}")
        return {}

def save_data(data):
    """Save application data to JSON file"""
    try:
        with open(DATA_FILE, 'w') as f:
            json.dump(data, f, indent=4)
        return True
    except Exception as e:
        logging.error(f"Error saving data: {str(e)}")
        return False

def parse_availability(raw_string):
    """Parse availability times from string format (e.g., 'Monday 12:00-15:00, Monday 20:00-00:00')"""
    if pd.isna(raw_string) or not raw_string:
        return {}
        
    day_map = {
        "sunday": "Sunday", "sun": "Sunday",
        "monday": "Monday", "mon": "Monday",
        "tuesday": "Tuesday", "tue": "Tuesday",
        "wednesday": "Wednesday", "wed": "Wednesday",
        "thursday": "Thursday", "thu": "Thursday",
        "friday": "Friday", "fri": "Friday",
        "saturday": "Saturday", "sat": "Saturday"
    }
    
    availability = {}
    
    # Split by commas and process each block
    blocks = re.split(r',\s*', str(raw_string))
    for block in blocks:
        # Match pattern like "Monday 12:00-15:00"
        match = re.match(r'(\w+)\s+(\d{1,2}:\d{2})-(\d{1,2}:\d{2})', block.strip(), re.IGNORECASE)
        if match:
            day_raw, start_time, end_time = match.groups()
            day_key = day_map.get(day_raw.lower(), None)
            
            if day_key:
                # Convert times to decimal hours for easier comparison
                start_hour = time_to_hour(start_time)
                end_hour = time_to_hour(end_time)
                
                # Handle overnight shifts (e.g., 22:00-02:00)
                if end_hour < start_hour:
                    end_hour += 24
                
                # Add to availability dictionary
                if day_key not in availability:
                    availability[day_key] = []
                
                availability[day_key].append({
                    "start": start_time,
                    "end": end_time,
                    "start_hour": start_hour,
                    "end_hour": end_hour
                })
    
    return availability

def time_to_hour(t):
    """Convert time string to decimal hour (e.g. '14:30' -> 14.5)"""
    if isinstance(t, str):
        parts = t.split(":")
        if len(parts) == 2:
            return int(parts[0]) + int(parts[1])/60
    return int(t)  # fallback if already an int

def hour_to_time_str(hour):
    """Convert decimal hour to time string (e.g. 14.5 -> '14:30')"""
    h = int(hour)
    m = int((hour - h) * 60)
    return f"{h:02d}:{m:02d}"

def format_time_ampm(time_str):
    """Format time string to AM/PM format"""
    try:
        hour, minute = map(int, time_str.split(':'))
        period = "AM" if hour < 12 else "PM"
        hour = hour % 12
        if hour == 0:
            hour = 12
        return f"{hour}:{minute:02d} {period}"
    except:
        return time_str

def overlaps(start1, end1, start2, end2):
    """Check if two time ranges overlap"""
    return max(start1, start2) < min(end1, end2)

def is_worker_available(worker, day, shift_start, shift_end):
    """Check if a worker is available for a given shift based on their availability"""
    # Get worker's availability for this day
    day_availability = worker.get('availability', {}).get(day, [])
    
    # If no specific availability is defined for this day, worker is not available
    if not day_availability:
        return False
    
    # Check if the shift overlaps with any of the worker's available times
    for avail in day_availability:
        avail_start = avail['start_hour']
        avail_end = avail['end_hour']
        
        # If the shift is completely within an available block, worker is available
        if avail_start <= shift_start and shift_end <= avail_end:
            return True
    
    # If we get here, no available block fully contains the shift
    return False

def find_alternative_workers(workers, day, shift_start, shift_end, assigned_hours, max_hours_per_worker, already_assigned):
    """Find alternative workers who could work this shift"""
    alternatives = []
    
    for worker in workers:
        email = worker['email']
        
        # Skip if already assigned to this shift
        if email in already_assigned:
            continue
            
        # Check if worker is available for this shift
        if is_worker_available(worker, day, shift_start, shift_end):
            # Check if adding this shift would exceed max hours
            shift_hours = shift_end - shift_start
            if assigned_hours.get(email, 0) + shift_hours <= max_hours_per_worker * 1.5: # Allow exceeding max hours for alternatives
                alternatives.append(worker)
    
    # Sort by assigned hours (least to most)
    alternatives.sort(key=lambda w: assigned_hours.get(w['email'], 0))
    
    return alternatives

def create_shifts_from_availability(hours_of_operation, workers, workplace, max_hours_per_worker, max_workers_per_shift):
    """Create shifts based on hours of operation and worker availability"""
    # Use timestamp as seed to ensure different schedules each time
    random.seed(datetime.now().timestamp())
    
    schedule = {}
    unfilled_shifts = []
    
    # Define possible shift lengths
    shift_lengths = [2, 3, 4, 5]  # 2, 3, 4, and 5 hour shifts
    
    # Randomize the order of shift lengths for more variety in schedules
    random.shuffle(shift_lengths)
    
    # track assigned hours per worker
    assigned_hours = {w['email']: 0 for w in workers}
    assigned_days = {w['email']: set() for w in workers}
    
    # track if a worker is work study (limited to exactly 5 hours per week)
    work_study_status = {w['email']: w.get('work_study', False) for w in workers}
    
    # Identify work study students who need exactly 5 hours
    work_study_workers = [w for w in workers if work_study_status[w['email']]]
    random.shuffle(work_study_workers)  # Randomize order for variety
    
    # First, try to assign 5-hour shifts to work study students
    for worker in work_study_workers:
        email = worker['email']
        
        # Skip if already assigned 5 hours
        if assigned_hours[email] >= 5:
            continue
            
        # Find a suitable 5-hour shift for this worker
        for day, operation_hours in hours_of_operation.items():
            if not operation_hours:
                continue
                
            for op in operation_hours:
                start_hour = time_to_hour(op['start'])
                end_hour = time_to_hour(op['end'])
                
                # Skip if invalid hours
                if end_hour <= start_hour:
                    end_hour += 24  # handle overnight shifts
                
                # Check if operation period is at least 5 hours
                if end_hour - start_hour >= 5:
                    # Get all possible 5-hour blocks and randomize them
                    possible_starts = [start_hour + i for i in range(int(end_hour - start_hour - 5) + 1)]
                    random.shuffle(possible_starts)
                    
                    # Try to find a 5-hour block where the worker is available
                    for potential_start in possible_starts:
                        potential_end = potential_start + 5
                        
                        if is_worker_available(worker, day, potential_start, potential_end):
                            # Initialize day in schedule if not exists
                            if day not in schedule:
                                schedule[day] = []
                            
                            # Add the 5-hour shift
                            schedule[day].append({
                                "start": hour_to_time_str(potential_start),
                                "end": hour_to_time_str(potential_end),
                                "assigned": [f"{worker['first_name']} {worker['last_name']}"],
                                "available": [f"{worker['first_name']} {worker['last_name']}"],
                                "raw_assigned": [email],
                                "all_available": [worker],
                                "is_work_study": True
                            })
                            
                            # Update assigned hours
                            assigned_hours[email] = 5
                            assigned_days[email].add(day)
                            
                            # Break once we've assigned a 5-hour shift
                            break
                    
                    # Break if we've assigned 5 hours
                    if assigned_hours[email] >= 5:
                        break
            
            # Break if we've assigned 5 hours
            if assigned_hours[email] >= 5:
                break
    
    # Now create regular shifts for the remaining time slots
    days_list = list(hours_of_operation.keys())
    random.shuffle(days_list)  # Randomize days for variety
    
    for day in days_list:
        operation_hours = hours_of_operation[day]
        if not operation_hours:
            continue  # skip days with no hours of operation
            
        if day not in schedule:
            schedule[day] = []
        
        # Randomize operation hours for variety
        random_operation_hours = operation_hours.copy()
        random.shuffle(random_operation_hours)
        
        # for each operation period in the day (e.g., morning and evening blocks)
        for op in random_operation_hours:
            start_hour = time_to_hour(op['start'])
            end_hour = time_to_hour(op['end'])
            
            # skip if invalid hours
            if end_hour <= start_hour:
                end_hour += 24  # handle overnight shifts
            
            # Check if there are any existing shifts for this day (from work study assignments)
            existing_shifts = [s for s in schedule[day] if 
                              overlaps(time_to_hour(s['start']), time_to_hour(s['end']), start_hour, end_hour)]
            
            # Create a list of time slots that need to be filled
            time_slots_to_fill = [(start_hour, end_hour)]
            
            # Remove time slots that are already covered by existing shifts
            for shift in existing_shifts:
                shift_start = time_to_hour(shift['start'])
                shift_end = time_to_hour(shift['end'])
                
                new_time_slots = []
                for slot_start, slot_end in time_slots_to_fill:
                    # If the slot is completely before or after the shift, keep it as is
                    if slot_end <= shift_start or slot_start >= shift_end:
                        new_time_slots.append((slot_start, slot_end))
                    else:
                        # If the slot overlaps with the shift, split it
                        if slot_start < shift_start:
                            new_time_slots.append((slot_start, shift_start))
                        if slot_end > shift_end:
                            new_time_slots.append((shift_end, slot_end))
                
                time_slots_to_fill = new_time_slots
            
            # For each remaining time slot, create shifts
            for slot_start, slot_end in time_slots_to_fill:
                slot_duration = slot_end - slot_start
                
                # Skip if slot is too short
                if slot_duration < 2:
                    continue
                
                # Determine which shift lengths to try for this slot
                possible_lengths = [l for l in shift_lengths if l <= slot_duration]
                if not possible_lengths:
                    possible_lengths = [2]  # Default to 2-hour shifts if nothing else fits
                
                # Randomize shift lengths for variety
                random.shuffle(possible_lengths)
                
                # Create shifts to cover the entire slot
                current_hour = slot_start
                while current_hour < slot_end:
                    # Randomize the shift length selection
                    # This makes the schedule different each time
                    random.shuffle(possible_lengths)
                    
                    # Find the best shift length that fits
                    shift_length = None
                    for length in possible_lengths:
                        if current_hour + length <= slot_end:
                            shift_length = length
                            break
                    
                    # If no shift length fits, use the smallest one and cap at slot_end
                    if shift_length is None:
                        shift_length = min(possible_lengths)
                    
                    shift_end_hour = min(current_hour + shift_length, slot_end)
                    
                    # find available workers for this shift
                    available_workers = []
                    for worker in workers:
                        email = worker['email']
                        
                        # Skip work study workers who already have their 5 hours
                        if work_study_status[email] and assigned_hours[email] >= 5:
                            continue
                        
                        # Skip work study workers for shifts that aren't 5 hours (unless they already have some hours)
                        if work_study_status[email] and assigned_hours[email] == 0 and (shift_end_hour - current_hour) != 5:
                            continue
                            
                        # check if worker is available
                        if is_worker_available(worker, day, current_hour, shift_end_hour):
                            # check max hours per worker limit
                            if assigned_hours.get(email, 0) + (shift_end_hour - current_hour) <= max_hours_per_worker:
                                # add to available workers
                                available_workers.append(worker)
                    
                    # Randomize the order of workers with the same hours
                    # This ensures different workers get assigned even with the same hours
                    available_workers.sort(key=lambda w: (assigned_hours[w['email']], random.random()))
                    
                    # assign workers to shift (up to max_workers_per_shift)
                    assigned = []
                    for worker in available_workers[:max_workers_per_shift]:
                        assigned.append(worker)
                        
                        # update worker's hours
                        email = worker['email']
                        assigned_hours[email] += (shift_end_hour - current_hour)
                        assigned_days[email].add(day)
                    
                    # Check if shift is unfilled
                    if not assigned:
                        unfilled_shifts.append({
                            "day": day,
                            "start": hour_to_time_str(current_hour),
                            "end": hour_to_time_str(shift_end_hour),
                            "start_hour": current_hour,
                            "end_hour": shift_end_hour
                        })
                    
                    # add shift to schedule
                    schedule[day].append({
                        "start": hour_to_time_str(current_hour),
                        "end": hour_to_time_str(shift_end_hour),
                        "assigned": [f"{w['first_name']} {w['last_name']}" for w in assigned] if assigned else ["Unfilled"],
                        "available": [f"{w['first_name']} {w['last_name']}" for w in available_workers],
                        "raw_assigned": [w['email'] for w in assigned] if assigned else [],
                        "all_available": [w for w in available_workers]  # store all available workers for editing
                    })
                    
                    # move to next shift
                    current_hour = shift_end_hour
    
    # identify workers with low hours
    low_hour_workers = []
    for w in workers:
        if not work_study_status.get(w['email'], False) and assigned_hours[w['email']] < 4:
            low_hour_workers.append(f"{w['first_name']} {w['last_name']}")
    
    # identify unassigned workers
    unassigned_workers = []
    for w in workers:
        if assigned_hours[w['email']] == 0:
            unassigned_workers.append(f"{w['first_name']} {w['last_name']}")
    
    # Check for work study students who didn't get exactly 5 hours
    work_study_issues = []
    for w in workers:
        if work_study_status.get(w['email'], False):
            hours = assigned_hours.get(w['email'], 0)
            if hours != 5:
                work_study_issues.append(f"{w['first_name']} {w['last_name']} ({hours} hours)")
    
    # Find alternative solutions for unfilled shifts
    alternative_solutions = {}
    for shift in unfilled_shifts:
        day = shift["day"]
        start_hour = shift["start_hour"]
        end_hour = shift["end_hour"]
        
        # Find workers who could work this shift if they worked more hours
        alternatives = find_alternative_workers(
            workers, 
            day, 
            start_hour, 
            end_hour, 
            {}, # Ignore current assigned hours to find all possibilities
            max_hours_per_worker * 1.5, # Allow exceeding max hours for alternatives
            []
        )
        
        if alternatives:
            alternative_solutions[f"{day} {shift['start']}-{shift['end']}"] = [
                f"{w['first_name']} {w['last_name']}" for w in alternatives
            ]
    
    return schedule, assigned_hours, low_hour_workers, unassigned_workers, alternative_solutions, unfilled_shifts, work_study_issues

def send_schedule_email(workplace, schedule, recipient_emails, sender_email, sender_password):
    """Send schedule via email"""
    try:
        # create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ", ".join(recipient_emails)
        msg['Subject'] = f"{workplace.replace('_', ' ').title()} Schedule"
        
        # create HTML body
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; }}
                table {{ border-collapse: collapse; width: 100%; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .unfilled {{ color: red; }}
            </style>
        </head>
        <body>
            <h2>{workplace.replace('_', ' ').title()} Schedule</h2>
        """
        
        # add schedule tables by day
        for day, shifts in schedule.items():
            if shifts:
                html += f"<h3>{day}</h3>"
                html += "<table>"
                html += "<tr><th>Start</th><th>End</th><th>Assigned</th></tr>"
                
                for shift in shifts:
                    assigned = ", ".join(shift['assigned'])
                    unfilled_class = ' class="unfilled"' if "Unfilled" in assigned else ""
                    
                    html += f"<tr>"
                    html += f"<td>{format_time_ampm(shift['start'])}</td>"
                    html += f"<td>{format_time_ampm(shift['end'])}</td>"
                    html += f"<td{unfilled_class}>{assigned}</td>"
                    html += f"</tr>"
                
                html += "</table>"
        
        html += """
        </body>
        </html>
        """
        
        # attach HTML body
        msg.attach(MIMEText(html, 'html'))
        
        # create schedule image
        img_path = create_schedule_image(workplace, schedule)
        if img_path and os.path.exists(img_path):
            with open(img_path, 'rb') as f:
                img = MIMEImage(f.read())
                img.add_header('Content-Disposition', 'attachment', filename=f"{workplace}_schedule.png")
                msg.attach(img)
        
        # create CSV file
        csv_path = create_schedule_csv(workplace, schedule)
        if csv_path and os.path.exists(csv_path):
            with open(csv_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype="csv")
                attachment.add_header('Content-Disposition', 'attachment', filename=f"{workplace}_schedule.csv")
                msg.attach(attachment)
        
        # create Excel file
        excel_path = create_schedule_excel(workplace, schedule)
        if excel_path and os.path.exists(excel_path):
            with open(excel_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype="xlsx")
                attachment.add_header('Content-Disposition', 'attachment', filename=f"{workplace}_schedule.xlsx")
                msg.attach(attachment)
        
        # send email
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        
        return True, "Email sent successfully"
    
    except Exception as e:
        logging.error(f"Error sending email: {str(e)}")
        return False, f"Error sending email: {str(e)}\n\nNote: For Gmail, you may need to use an App Password instead of your regular password. Go to your Google Account > Security > App Passwords to create one."

def create_schedule_image(workplace, schedule):
    """Create an image of the schedule"""
    try:
        # flatten schedule into rows
        rows = []
        for day, shifts in schedule.items():
            for shift in shifts:
                rows.append({
                    "Day": day,
                    "Start": shift['start'],
                    "End": shift['end'],
                    "Assigned": ", ".join(shift['assigned'])
                })
        
        if not rows:
            return None
        
        # create figure and axis
        fig, ax = plt.subplots(figsize=(10, len(rows) * 0.4))
        ax.axis('off')
        
        # create table data
        table_data = [["Day", "Start", "End", "Assigned"]] + [[r["Day"], format_time_ampm(r["Start"]), format_time_ampm(r["End"]), r["Assigned"]] for r in rows]
        
        # create table
        table = ax.table(cellText=table_data, cellLoc='center', loc='center')
        
        # style table
        for cell in table.get_celld().values():
            cell.set_fontsize(10)
        
        # save figure
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(DIRS['schedules'], f"{workplace}_{timestamp}.png")
        plt.savefig(output_path, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    except Exception as e:
        logging.error(f"Error creating schedule image: {str(e)}")
        return None

def create_schedule_csv(workplace, schedule):
    """Create a CSV file of the schedule"""
    try:
        # flatten schedule into rows
        rows = []
        for day, shifts in schedule.items():
            for shift in shifts:
                rows.append({
                    "Day": day,
                    "Start": format_time_ampm(shift['start']),
                    "End": format_time_ampm(shift['end']),
                    "Assigned": ", ".join(shift['assigned'])
                })
        
        if not rows:
            return None
        
        # create DataFrame
        df = pd.DataFrame(rows)
        
        # save to CSV
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(DIRS['schedules'], f"{workplace}_{timestamp}.csv")
        df.to_csv(output_path, index=False)
        
        return output_path
    
    except Exception as e:
        logging.error(f"Error creating schedule CSV: {str(e)}")
        return None

def create_schedule_excel(workplace, schedule):
    """Create an Excel file of the schedule"""
    try:
        # Create a DataFrame for each day
        dfs = {}
        for day in DAYS:
            if day in schedule and schedule[day]:
                day_shifts = schedule[day]
                rows = []
                for shift in day_shifts:
                    rows.append({
                        "Start": format_time_ampm(shift['start']),
                        "End": format_time_ampm(shift['end']),
                        "Assigned": ", ".join(shift['assigned'])
                    })
                dfs[day] = pd.DataFrame(rows)
        
        if not dfs:
            return None
        
        # Create a writer for Excel file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(DIRS['schedules'], f"{workplace}_{timestamp}.xlsx")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write each day to a separate sheet
            for day, df in dfs.items():
                df.to_excel(writer, sheet_name=day, index=False)
            
            # Create a summary sheet
            all_shifts = []
            for day, shifts in schedule.items():
                for shift in shifts:
                    all_shifts.append({
                        "Day": day,
                        "Start": format_time_ampm(shift['start']),
                        "End": format_time_ampm(shift['end']),
                        "Assigned": ", ".join(shift['assigned'])
                    })
            
            if all_shifts:
                summary_df = pd.DataFrame(all_shifts)
                summary_df.to_excel(writer, sheet_name="Full Schedule", index=False)
        
        return output_path
    
    except Exception as e:
        logging.error(f"Error creating schedule Excel: {str(e)}")
        return None

def find_available_workers(workers, day, start_time, end_time):
    """Find workers available for a specific time slot"""
    available_workers = []
    
    start_hour = time_to_hour(start_time)
    end_hour = time_to_hour(end_time)
    
    for worker in workers:
        if is_worker_available(worker, day, start_hour, end_hour):
            available_workers.append(worker)
    
    return available_workers

# main application classes
class StyleHelper:
    """Helper class for consistent styling"""
    
    @staticmethod
    def get_main_style():
        return """
            QMainWindow, QDialog {
                background-color: #f0f2f5;
            }
            QTabWidget::pane {
                border: 1px solid #ddd;
                background-color: white;
                border-radius: 5px;
            }
            QTabBar::tab {
                background-color: #e0e0e0;
                border: 1px solid #ccc;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                padding: 8px 12px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: 1px solid white;
            }
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004494;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
            QLabel {
                color: #333;
            }
            QLineEdit, QComboBox, QSpinBox, QTimeEdit, QDateEdit {
                padding: 6px;
                border: 1px solid #ccc;
                border-radius: 4px;
            }
            QTableWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
            }
            QTableWidget::item {
                padding: 4px;
            }
            QTableWidget::item:selected {
                background-color: #e7f0fd;
                color: black;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 6px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QGroupBox {
                border: 1px solid #ddd;
                border-radius: 5px;
                margin-top: 10px;
                font-weight: bold;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """
    
    @staticmethod
    def create_section_title(text):
        label = QLabel(text)
        font = label.font()
        font.setPointSize(12)
        font.setBold(True)
        label.setFont(font)
        return label
    
    @staticmethod
    def create_button(text, primary=True):
        btn = QPushButton(text)
        if not primary:
            btn.setStyleSheet("""
                background-color: #6c757d;
                color: white;
            """)
        return btn
    
    @staticmethod
    def create_action_button(text):
        btn = QPushButton(text)
        btn.setStyleSheet("""
            background-color: #28a745;
            color: white;
        """)
        return btn
    
    @staticmethod
    def create_warning_button(text):
        btn = QPushButton(text)
        btn.setStyleSheet("""
            background-color: #dc3545;
            color: white;
        """)
        return btn

class DayTimeBlockWidget(QWidget):
    """Widget for managing a single day's time blocks"""
    
    def __init__(self, day, parent=None):
        super().__init__(parent)
        self.day = day
        self.time_blocks = []
        self.initUI()
    
    def initUI(self):
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        
        # Day label
        day_label = QLabel(self.day)
        day_label.setStyleSheet("font-weight: bold;")
        self.layout.addWidget(day_label)
        
        # Container for time blocks
        self.blocks_container = QWidget()
        self.blocks_layout = QVBoxLayout(self.blocks_container)
        self.blocks_layout.setContentsMargins(0, 0, 0, 0)
        self.layout.addWidget(self.blocks_container)
        
        # Add button
        add_btn = QPushButton("Add Time Block")
        add_btn.clicked.connect(self.add_time_block)
        self.layout.addWidget(add_btn)
    
    def add_time_block(self):
        """Add a new time block"""
        block_widget = QWidget()
        block_layout = QHBoxLayout(block_widget)
        block_layout.setContentsMargins(0, 0, 0, 0)
        
        # Start time
        start_time = QTimeEdit()
        start_time.setDisplayFormat("HH:mm")
        start_time.setTime(QTime(9, 0))
        
        # End time
        end_time = QTimeEdit()
        end_time.setDisplayFormat("HH:mm")
        end_time.setTime(QTime(17, 0))
        
        # Remove button
        remove_btn = QPushButton("Remove")
        remove_btn.setStyleSheet("background-color: #dc3545;")
        remove_btn.clicked.connect(lambda: self.remove_time_block(block_widget))
        
        block_layout.addWidget(QLabel("Start:"))
        block_layout.addWidget(start_time)
        block_layout.addWidget(QLabel("End:"))
        block_layout.addWidget(end_time)
        block_layout.addWidget(remove_btn)
        
        self.blocks_layout.addWidget(block_widget)
        self.time_blocks.append((start_time, end_time))
    
    def remove_time_block(self, block_widget):
        """Remove a time block"""
        index = -1
        for i in range(self.blocks_layout.count()):
            if self.blocks_layout.itemAt(i).widget() == block_widget:
                index = i
                break
        
        if index >= 0:
            self.blocks_layout.itemAt(index).widget().deleteLater()
            self.time_blocks.pop(index)
    
    def set_blocks(self, blocks):
        """Set time blocks from data"""
        # Clear existing blocks
        while self.blocks_layout.count():
            widget = self.blocks_layout.itemAt(0).widget()
            if widget:
                widget.deleteLater()
        self.time_blocks = []
        
        # Add blocks from data
        for block in blocks:
            self.add_time_block_with_data(block)
    
    def add_time_block_with_data(self, block):
        """Add a time block with specific data"""
        block_widget = QWidget()
        block_layout = QHBoxLayout(block_widget)
        block_layout.setContentsMargins(0, 0, 0, 0)
        
        # Start time
        start_time = QTimeEdit()
        start_time.setDisplayFormat("HH:mm")
        if 'start' in block:
            start_time.setTime(QTime.fromString(block['start'], "HH:mm"))
        
        # End time
        end_time = QTimeEdit()
        end_time.setDisplayFormat("HH:mm")
        if 'end' in block:
            end_time.setTime(QTime.fromString(block['end'], "HH:mm"))
        
        # Remove button
        remove_btn = QPushButton("Remove")
        remove_btn.setStyleSheet("background-color: #dc3545;")
        remove_btn.clicked.connect(lambda: self.remove_time_block(block_widget))
        
        block_layout.addWidget(QLabel("Start:"))
        block_layout.addWidget(start_time)
        block_layout.addWidget(QLabel("End:"))
        block_layout.addWidget(end_time)
        block_layout.addWidget(remove_btn)
        
        self.blocks_layout.addWidget(block_widget)
        self.time_blocks.append((start_time, end_time))
    
    def get_blocks(self):
        """Get time blocks as data"""
        blocks = []
        for start_time, end_time in self.time_blocks:
            blocks.append({
                "start": start_time.time().toString("HH:mm"),
                "end": end_time.time().toString("HH:mm")
            })
        return blocks

class HoursOfOperationDialog(QDialog):
    """Dialog for managing hours of operation"""
    
    def __init__(self, workplace, hours_data, parent=None):
        super().__init__(parent)
        self.workplace = workplace
        self.hours_data = hours_data
        self.day_widgets = {}
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle(f"Hours of Operation - {self.workplace.replace('_', ' ').title()}")
        self.setMinimumWidth(500)
        self.setMinimumHeight(600)
        
        layout = QVBoxLayout()
        
        # Scroll area for days
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # Create widgets for each day
        for day in DAYS:
            day_widget = DayTimeBlockWidget(day)
            blocks = self.hours_data.get(day, [])
            day_widget.set_blocks(blocks)
            
            scroll_layout.addWidget(day_widget)
            self.day_widgets[day] = day_widget
        
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        
        layout.addWidget(scroll_area)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        save_btn = StyleHelper.create_button("Save")
        save_btn.clicked.connect(self.save_hours)
        
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        cancel_btn.clicked.connect(self.reject)
        
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
    
    def save_hours(self):
        """Save hours of operation"""
        hours_data = {}
        
        for day, widget in self.day_widgets.items():
            hours_data[day] = widget.get_blocks()
        
        self.hours_data = hours_data
        self.accept()

class AlternativeSolutionsDialog(QDialog):
    """Dialog for showing alternative solutions for unfilled shifts"""
    
    def __init__(self, alternative_solutions, unfilled_shifts, work_study_issues=None, parent=None):
        super().__init__(parent)
        self.alternative_solutions = alternative_solutions
        self.unfilled_shifts = unfilled_shifts
        self.work_study_issues = work_study_issues or []
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle("Schedule Suggestions")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Suggestions for Unfilled Shifts")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title)
        
        # Explanation
        explanation = QLabel("The following shifts are currently unfilled. Here are some suggestions to fill them:")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)
        
        # Create a scroll area for the content
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # Work Study Issues
        if self.work_study_issues:
            ws_group = QGroupBox("Work Study Issues")
            ws_group.setStyleSheet("background-color: #fff3cd; border: 1px solid #ffeeba; border-radius: 5px;")
            ws_layout = QVBoxLayout()
            
            ws_label = QLabel("The following work study students don't have exactly 5 hours:")
            ws_label.setWordWrap(True)
            ws_layout.addWidget(ws_label)
            
            for worker in self.work_study_issues:
                worker_label = QLabel(f"• {worker}")
                worker_label.setStyleSheet("font-weight: bold;")
                ws_layout.addWidget(worker_label)
            
            suggestion = QLabel("Suggestion: Work study students must have exactly 5 hours per week. Try adjusting their shifts manually.")
            suggestion.setStyleSheet("font-style: italic;")
            suggestion.setWordWrap(True)
            ws_layout.addWidget(suggestion)
            
            ws_group.setLayout(ws_layout)
            scroll_layout.addWidget(ws_group)
        
        # Check if there are unfilled shifts
        if not self.unfilled_shifts:
            no_unfilled = QLabel("All shifts are filled! Great job!")
            no_unfilled.setStyleSheet("font-weight: bold; color: green;")
            scroll_layout.addWidget(no_unfilled)
        else:
            # For each unfilled shift
            for shift in self.unfilled_shifts:
                day = shift["day"]
                start = format_time_ampm(shift["start"])
                end = format_time_ampm(shift["end"])
                
                shift_key = f"{day} {shift['start']}-{shift['end']}"
                
                # Create a group box for each shift
                shift_group = QGroupBox(f"{day} {start} - {end}")
                shift_group.setStyleSheet("background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 5px;")
                shift_layout = QVBoxLayout()
                
                # If we have alternative solutions
                if shift_key in self.alternative_solutions and self.alternative_solutions[shift_key]:
                    alternatives = self.alternative_solutions[shift_key]
                    
                    # Add a label explaining the alternatives
                    alt_label = QLabel(f"The following workers are available but would exceed their hour limits:")
                    alt_label.setWordWrap(True)
                    shift_layout.addWidget(alt_label)
                    
                    # Add each alternative worker
                    for worker in alternatives:
                        worker_label = QLabel(f"• {worker}")
                        worker_label.setStyleSheet("font-weight: bold;")
                        shift_layout.addWidget(worker_label)
                    
                    # Add suggestion
                    suggestion = QLabel("Suggestion: Consider increasing their max hours or reassigning other shifts.")
                    suggestion.setStyleSheet("font-style: italic;")
                    suggestion.setWordWrap(True)
                    shift_layout.addWidget(suggestion)
                else:
                    # No alternatives available
                    no_alt_label = QLabel("No workers are available for this shift, even with extended hours.")
                    no_alt_label.setWordWrap(True)
                    shift_layout.addWidget(no_alt_label)
                    
                    suggestion = QLabel("Suggestion: Consider adjusting hours of operation or recruiting more workers with availability during this time.")
                    suggestion.setStyleSheet("font-style: italic;")
                    suggestion.setWordWrap(True)
                    shift_layout.addWidget(suggestion)
                
                shift_group.setLayout(shift_layout)
                scroll_layout.addWidget(shift_group)
        
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)
        
        # Close button
        close_btn = StyleHelper.create_button("Close", primary=False)
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)

class LastMinuteAvailabilityDialog(QDialog):
    """Dialog for checking last minute availability"""
    
    def __init__(self, workplace, parent=None):
        super().__init__(parent)
        self.workplace = workplace
        self.workers = []
        self.initUI()
        self.loadWorkers()
    
    def initUI(self):
        self.setWindowTitle(f"Last Minute Availability - {self.workplace.replace('_', ' ').title()}")
        self.setMinimumWidth(700)
        self.setMinimumHeight(500)
        
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Check Last Minute Availability")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title)
        
        # Form for selecting day and time
        form_layout = QFormLayout()
        
        # Day selection
        self.day_combo = QComboBox()
        self.day_combo.addItems(DAYS)
        form_layout.addRow("Day:", self.day_combo)
        
        # Time selection
        time_layout = QHBoxLayout()
        
        self.start_time = QTimeEdit()
        self.start_time.setDisplayFormat("HH:mm")
        self.start_time.setTime(QTime(9, 0))
        
        self.end_time = QTimeEdit()
        self.end_time.setDisplayFormat("HH:mm")
        self.end_time.setTime(QTime(17, 0))
        
        time_layout.addWidget(QLabel("Start:"))
        time_layout.addWidget(self.start_time)
        time_layout.addWidget(QLabel("End:"))
        time_layout.addWidget(self.end_time)
        
        form_layout.addRow("Time:", time_layout)
        
        layout.addLayout(form_layout)
        
        # Check button
        check_btn = StyleHelper.create_action_button("Check Availability")
        check_btn.clicked.connect(self.checkAvailability)
        layout.addWidget(check_btn)
        
        # Results table
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(3)
        self.results_table.setHorizontalHeaderLabels(["Name", "Email", "Work Study"])
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.results_table)
        
        # Close button
        close_btn = StyleHelper.create_button("Close", primary=False)
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)
    
    def loadWorkers(self):
        """Load workers from Excel file"""
        file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
        
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Warning", "No Excel file found for this workplace.")
            return
        
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Clean the DataFrame
            df = df.dropna(subset=['Email'], how='all')
            df = df[df['Email'].str.strip() != '']
            df = df[~df['Email'].str.contains('nan', case=False, na=False)]
            
            self.workers = []
            for _, row in df.iterrows():
                # Get availability from the "Days & Times Available" column
                avail_column = next((col for col in df.columns if 'available' in col.lower()), None)
                availability_text = str(row.get(avail_column, "")) if avail_column else ""
                if pd.isna(availability_text) or availability_text == "nan":
                    availability_text = ""
                
                # Parse availability into structured format
                availability = parse_availability(availability_text)
                
                self.workers.append({
                    "first_name": row.get("First Name", "").strip(),
                    "last_name": row.get("Last Name", "").strip(),
                    "email": row.get("Email", "").strip(),
                    "work_study": str(row.get("Work Study", "")).strip().lower() in ['yes', 'y', 'true'],
                    "availability": availability
                })
            
        except Exception as e:
            logging.error(f"Error loading workers: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error loading workers: {str(e)}")
    
    def checkAvailability(self):
        """Check which workers are available for the selected time"""
        day = self.day_combo.currentText()
        start_time = self.start_time.time().toString("HH:mm")
        end_time = self.end_time.time().toString("HH:mm")
        
        # Find available workers
        available_workers = find_available_workers(self.workers, day, start_time, end_time)
        
        # Display results
        self.results_table.setRowCount(len(available_workers))
        
        for i, worker in enumerate(available_workers):
            # Name
            name = f"{worker['first_name']} {worker['last_name']}"
            self.results_table.setItem(i, 0, QTableWidgetItem(name))
            
            # Email
            self.results_table.setItem(i, 1, QTableWidgetItem(worker['email']))
            
            # Work Study
            work_study = "Yes" if worker['work_study'] else "No"
            self.results_table.setItem(i, 2, QTableWidgetItem(work_study))
        
        # Show message if no workers are available
        if not available_workers:
            QMessageBox.warning(self, "No Available Workers", 
                               f"No workers are available on {day} from {format_time_ampm(start_time)} to {format_time_ampm(end_time)}.")

class WorkplaceTab(QWidget):
    """Tab for managing a specific workplace"""
    
    def __init__(self, workplace, parent=None):
        super().__init__(parent)
        self.workplace = workplace
        self.app_data = load_data()
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout()
        
        # workplace title
        title = StyleHelper.create_section_title(f"{self.workplace.replace('_', ' ').title()}")
        layout.addWidget(title)
        
        # quick actions
        actions_layout = QHBoxLayout()
        
        upload_btn = StyleHelper.create_button("Upload Excel File")
        upload_btn.clicked.connect(self.upload_excel)
        
        hours_btn = StyleHelper.create_button("Hours of Operation")
        hours_btn.clicked.connect(self.manage_hours)
        
        generate_btn = StyleHelper.create_action_button("Generate Schedule")
        generate_btn.clicked.connect(self.generate_schedule)
        
        view_btn = StyleHelper.create_button("View Current Schedule", primary=False)
        view_btn.clicked.connect(self.view_current_schedule)
        
        last_minute_btn = StyleHelper.create_button("Last Minute", primary=False)
        last_minute_btn.setStyleSheet("background-color: #fd7e14; color: white;")
        last_minute_btn.clicked.connect(self.show_last_minute_dialog)
        
        actions_layout.addWidget(upload_btn)
        actions_layout.addWidget(hours_btn)
        actions_layout.addWidget(generate_btn)
        actions_layout.addWidget(view_btn)
        actions_layout.addWidget(last_minute_btn)
        
        layout.addLayout(actions_layout)
        
        # tab widget for different sections
        self.tabs = QTabWidget()
        
        # workers tab
        workers_tab = QWidget()
        workers_layout = QVBoxLayout()
        
        # workers table
        self.workers_table = QTableWidget()
        self.workers_table.setColumnCount(6)
        self.workers_table.setHorizontalHeaderLabels(["First Name", "Last Name", "Email", "Work Study", "Availability", "Actions"])
        
        # load workers
        self.load_workers_table(self.workers_table)
        
        workers_layout.addWidget(self.workers_table)
        
        # add worker button
        add_worker_btn = StyleHelper.create_button("Add Worker")
        add_worker_btn.clicked.connect(lambda: self.add_worker_dialog(self.workers_table))
        workers_layout.addWidget(add_worker_btn)
        
        workers_tab.setLayout(workers_layout)
        
        # hours of operation tab
        hours_tab = QWidget()
        hours_layout = QVBoxLayout()
        
        hours_group = QGroupBox("Current Hours of Operation")
        hours_group_layout = QVBoxLayout()
        
        # load hours of operation
        self.hours_table = QTableWidget()
        self.hours_table.setColumnCount(3)
        self.hours_table.setHorizontalHeaderLabels(["Day", "Start", "End"])
        
        self.load_hours_table(self.hours_table)
        
        hours_group_layout.addWidget(self.hours_table)
        hours_group.setLayout(hours_group_layout)
        hours_layout.addWidget(hours_group)
        
        edit_hours_btn = StyleHelper.create_button("Edit Hours of Operation")
        edit_hours_btn.clicked.connect(self.manage_hours)
        hours_layout.addWidget(edit_hours_btn)
        
        hours_tab.setLayout(hours_layout)
        
        # last minute tab
        last_minute_tab = QWidget()
        last_minute_layout = QVBoxLayout()
        
        last_minute_title = QLabel("Last Minute Availability")
        last_minute_title.setStyleSheet("font-size: 14px; font-weight: bold;")
        last_minute_layout.addWidget(last_minute_title)
        
        last_minute_desc = QLabel("Check which workers are available for a specific time slot.")
        last_minute_desc.setWordWrap(True)
        last_minute_layout.addWidget(last_minute_desc)
        
        # Form for selecting day and time
        form_layout = QFormLayout()
        
        # Day selection
        self.lm_day_combo = QComboBox()
        self.lm_day_combo.addItems(DAYS)
        form_layout.addRow("Day:", self.lm_day_combo)
        
        # Time selection
        time_layout = QHBoxLayout()
        
        self.lm_start_time = QTimeEdit()
        self.lm_start_time.setDisplayFormat("HH:mm")
        self.lm_start_time.setTime(QTime(9, 0))
        
        self.lm_end_time = QTimeEdit()
        self.lm_end_time.setDisplayFormat("HH:mm")
        self.lm_end_time.setTime(QTime(17, 0))
        
        time_layout.addWidget(QLabel("Start:"))
        time_layout.addWidget(self.lm_start_time)
        time_layout.addWidget(QLabel("End:"))
        time_layout.addWidget(self.lm_end_time)
        
        form_layout.addRow("Time:", time_layout)
        
        last_minute_layout.addLayout(form_layout)
        
        # Check button
        check_btn = StyleHelper.create_action_button("Check Availability")
        check_btn.clicked.connect(self.check_last_minute_availability)
        last_minute_layout.addWidget(check_btn)
        
        # Results table
        self.lm_results_table = QTableWidget()
        self.lm_results_table.setColumnCount(3)
        self.lm_results_table.setHorizontalHeaderLabels(["Name", "Email", "Work Study"])
        self.lm_results_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        last_minute_layout.addWidget(self.lm_results_table)
        
        last_minute_tab.setLayout(last_minute_layout)
        
        # add tabs
        self.tabs.addTab(workers_tab, "Workers")
        self.tabs.addTab(hours_tab, "Hours of Operation")
        self.tabs.addTab(last_minute_tab, "Last Minute")
        
        layout.addWidget(self.tabs)
        
        self.setLayout(layout)
    
    def load_workers_table(self, table):
        """Load workers into table"""
        # clear table
        table.setRowCount(0)
        
        # check if Excel file exists
        file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
        if not os.path.exists(file_path):
            return
        
        # load Excel file
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Filter out rows that don't have valid data
            df = df.dropna(subset=['Email'], how='all')
            df = df[df['Email'].str.strip() != '']
            df = df[~df['Email'].str.contains('nan', case=False, na=False)]
            
            # set row count
            table.setRowCount(len(df))
            
            # fill table
            for i, (_, row) in enumerate(df.iterrows()):
                # first name
                first_name = row.get("First Name", "")
                if pd.isna(first_name) or first_name == "nan":
                    first_name = ""
                table.setItem(i, 0, QTableWidgetItem(str(first_name)))
                
                # last name
                last_name = row.get("Last Name", "")
                if pd.isna(last_name) or last_name == "nan":
                    last_name = ""
                table.setItem(i, 1, QTableWidgetItem(str(last_name)))
                
                # email
                email = row.get("Email", "")
                if pd.isna(email) or email == "nan":
                    email = ""
                table.setItem(i, 2, QTableWidgetItem(str(email)))
                
                # work study
                work_study = row.get("Work Study", "No")
                if pd.isna(work_study) or work_study == "nan":
                    work_study = "No"
                table.setItem(i, 3, QTableWidgetItem(str(work_study)))
                
                # availability
                avail_column = next((col for col in df.columns if 'available' in col.lower()), None)
                availability_text = str(row.get(avail_column, "")) if avail_column else ""
                if pd.isna(availability_text) or availability_text == "nan":
                    availability_text = ""
                table.setItem(i, 4, QTableWidgetItem(availability_text))
                
                # actions
                actions_widget = QWidget()
                actions_layout = QHBoxLayout()
                actions_layout.setContentsMargins(0, 0, 0, 0)
                
                edit_btn = QPushButton("Edit")
                edit_btn.setStyleSheet("background-color: #ffc107; color: black;")
                edit_btn.clicked.connect(lambda _, r=i, e=email: self.edit_worker_dialog(table, r, e))
                
                delete_btn = QPushButton("Delete")
                delete_btn.setStyleSheet("background-color: #dc3545;")
                delete_btn.clicked.connect(lambda _, e=email: self.delete_worker(table, e))
                
                actions_layout.addWidget(edit_btn)
                actions_layout.addWidget(delete_btn)
                
                actions_widget.setLayout(actions_layout)
                table.setCellWidget(i, 5, actions_widget)
            
            # resize columns
            table.resizeColumnsToContents()
            
            # make sure we're on the Workers tab after loading
            self.tabs.setCurrentIndex(0)
            
        except Exception as e:
            logging.error(f"Error loading workers: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error loading workers: {str(e)}")
    
    def load_hours_table(self, table):
        """Load hours of operation into table"""
        # clear table
        table.setRowCount(0)
        
        # get hours of operation
        hours = {}
        if self.workplace in self.app_data and 'hours_of_operation' in self.app_data[self.workplace]:
            hours = self.app_data[self.workplace]['hours_of_operation']
        
        # count total rows needed
        total_rows = sum(len(blocks) if blocks else 1 for blocks in hours.values())
        table.setRowCount(total_rows)
        
        # fill table
        row_index = 0
        for day in DAYS:
            blocks = hours.get(day, [])
            
            if not blocks:
                # no hours for this day
                table.setItem(row_index, 0, QTableWidgetItem(day))
                table.setItem(row_index, 1, QTableWidgetItem("Closed"))
                table.setItem(row_index, 2, QTableWidgetItem("Closed"))
                row_index += 1
            else:
                # hours for this day
                for block in blocks:
                    table.setItem(row_index, 0, QTableWidgetItem(day))
                    table.setItem(row_index, 1, QTableWidgetItem(format_time_ampm(block['start'])))
                    table.setItem(row_index, 2, QTableWidgetItem(format_time_ampm(block['end'])))
                    row_index += 1
        
        # resize columns
        table.resizeColumnsToContents()
    
    def upload_excel(self):
        """Upload Excel file for workplace"""
        file_path, _ = QFileDialog.getOpenFileName(self, "Upload Excel File", "", "Excel Files (*.xlsx)")
        
        if not file_path:
            return
        
        try:
            # copy file to workplaces directory
            destination = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
            import shutil
            shutil.copy2(file_path, destination)
            
            # Clean up the Excel file
            self.clean_excel_file(destination)
            
            # reload workers table
            self.load_workers_table(self.workers_table)
            
            QMessageBox.information(self, "Success", "Excel file uploaded successfully.")
            
        except Exception as e:
            logging.error(f"Error uploading Excel file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error uploading Excel file: {str(e)}")
    
    def clean_excel_file(self, file_path):
        """Clean up the Excel file to remove empty rows and fix formatting"""
        try:
            # Read the Excel file
            df = pd.read_excel(file_path)
            
            # Clean column names
            df.columns = df.columns.str.strip()
            
            # Filter out rows with empty or 'nan' emails
            df = df.dropna(subset=['Email'], how='all')
            df = df[df['Email'].str.strip() != '']
            df = df[~df['Email'].str.contains('nan', case=False, na=False)]
            
            # Save the cleaned file
            df.to_excel(file_path, index=False)
            
        except Exception as e:
            logging.error(f"Error cleaning Excel file: {str(e)}")
            raise
    
    def add_worker_dialog(self, table):
        """Show dialog to add a worker"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Worker")
        dialog.setMinimumWidth(500)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # first name
        first_name_input = QLineEdit()
        form_layout.addRow("First Name:", first_name_input)
        
        # last name
        last_name_input = QLineEdit()
        form_layout.addRow("Last Name:", last_name_input)
        
        # email
        email_input = QLineEdit()
        form_layout.addRow("Email:", email_input)
        
        # work study
        work_study_combo = QComboBox()
        work_study_combo.addItems(["No", "Yes"])
        form_layout.addRow("Work Study:", work_study_combo)
        
        # availability
        avail_input = QTextEdit()
        avail_input.setPlaceholderText("Enter availability in format: Day HH:MM-HH:MM\nExample: Monday 12:00-15:00, Monday 20:00-00:00, Tuesday 12:00-15:00")
        avail_input.setMinimumHeight(100)
        form_layout.addRow("Days & Times Available:", avail_input)
        
        layout.addLayout(form_layout)
        
        # buttons
        buttons_layout = QHBoxLayout()
        
        save_btn = StyleHelper.create_button("Save")
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # connect buttons
        save_btn.clicked.connect(lambda: self.save_worker(
            dialog,
            table,
            first_name_input.text(),
            last_name_input.text(),
            email_input.text(),
            work_study_combo.currentText(),
            avail_input.toPlainText()
        ))
        
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec_()
    
    def save_worker(self, dialog, table, first_name, last_name, email, work_study, availability):
        """Save worker to Excel file"""
        if not first_name or not last_name or not email:
            QMessageBox.warning(dialog, "Warning", "First name, last name, and email are required.")
            return
        
        try:
            # check if Excel file exists
            file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
            
            if os.path.exists(file_path):
                # load existing file
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                
                # Clean the DataFrame
                df = df.dropna(subset=['Email'], how='all')
                df = df[df['Email'].str.strip() != '']
                df = df[~df['Email'].str.contains('nan', case=False, na=False)]
                
                # check if email already exists
                if email in df['Email'].values:
                    QMessageBox.warning(dialog, "Warning", "A worker with this email already exists.")
                    return
                
                # create new row
                new_row = {
                    "First Name": first_name,
                    "Last Name": last_name,
                    "Email": email,
                    "Work Study": work_study
                }
                
                # add availability
                avail_column = next((col for col in df.columns if 'available' in col.lower()), None)
                if avail_column:
                    new_row[avail_column] = availability
                
                # append row
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                
            else:
                # create new file
                columns = ["First Name", "Last Name", "Email", "Work Study", "Days & Times Available"]
                df = pd.DataFrame(columns=columns)
                
                # create new row
                new_row = {
                    "First Name": first_name,
                    "Last Name": last_name,
                    "Email": email,
                    "Work Study": work_study,
                    "Days & Times Available": availability
                }
                
                # append row
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            # save file
            df.to_excel(file_path, index=False)
            
            # reload workers table
            self.load_workers_table(table)
            
            dialog.accept()
            
        except Exception as e:
            logging.error(f"Error saving worker: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error saving worker: {str(e)}")
    
    def edit_worker_dialog(self, table, row, email):
        """Show dialog to edit a worker"""
        # get worker data
        file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
        
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Warning", "Excel file not found.")
            return
        
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # find worker
            worker_row = df[df['Email'] == email]
            
            if worker_row.empty:
                QMessageBox.warning(self, "Warning", "Worker not found.")
                return
            
            worker_row = worker_row.iloc[0]
            
            # create dialog
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Worker")
            dialog.setMinimumWidth(500)
            
            layout = QVBoxLayout()
            
            form_layout = QFormLayout()
            
            # first name
            first_name_input = QLineEdit()
            first_name_input.setText(str(worker_row.get("First Name", "")))
            form_layout.addRow("First Name:", first_name_input)
            
            # last name
            last_name_input = QLineEdit()
            last_name_input.setText(str(worker_row.get("Last Name", "")))
            form_layout.addRow("Last Name:", last_name_input)
            
            # email
            email_input = QLineEdit()
            email_input.setText(str(worker_row.get("Email", "")))
            email_input.setReadOnly(True)  # email cannot be changed
            form_layout.addRow("Email:", email_input)
            
            # work study
            work_study_combo = QComboBox()
            work_study_combo.addItems(["No", "Yes"])
            work_study_combo.setCurrentText(str(worker_row.get("Work Study", "No")))
            form_layout.addRow("Work Study:", work_study_combo)
            
            # availability
            avail_column = next((col for col in df.columns if 'available' in col.lower()), None)
            avail_input = QTextEdit()
            if avail_column:
                avail_input.setText(str(worker_row.get(avail_column, "")))
            avail_input.setPlaceholderText("Enter availability in format: Day HH:MM-HH:MM\nExample: Monday 12:00-15:00, Monday 20:00-00:00, Tuesday 12:00-15:00")
            avail_input.setMinimumHeight(100)
            form_layout.addRow("Days & Times Available:", avail_input)
            
            layout.addLayout(form_layout)
            
            # buttons
            buttons_layout = QHBoxLayout()
            
            save_btn = StyleHelper.create_button("Save")
            cancel_btn = StyleHelper.create_button("Cancel", primary=False)
            
            buttons_layout.addWidget(save_btn)
            buttons_layout.addWidget(cancel_btn)
            
            layout.addLayout(buttons_layout)
            
            dialog.setLayout(layout)
            
            # connect buttons
            save_btn.clicked.connect(lambda: self.update_worker(
                dialog,
                table,
                email,
                first_name_input.text(),
                last_name_input.text(),
                work_study_combo.currentText(),
                avail_input.toPlainText()
            ))
            
            cancel_btn.clicked.connect(dialog.reject)
            
            dialog.exec_()
            
        except Exception as e:
            logging.error(f"Error editing worker: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error editing worker: {str(e)}")
    
    def update_worker(self, dialog, table, email, first_name, last_name, work_study, availability):
        """Update worker in Excel file"""
        if not first_name or not last_name:
            QMessageBox.warning(dialog, "Warning", "First name and last name are required.")
            return
        
        try:
            # load Excel file
            file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # find worker
            mask = df['Email'] == email
            
            if not any(mask):
                QMessageBox.warning(dialog, "Warning", "Worker not found.")
                return
            
            # update worker
            df.loc[mask, "First Name"] = first_name
            df.loc[mask, "Last Name"] = last_name
            df.loc[mask, "Work Study"] = work_study
            
            # update availability
            avail_column = next((col for col in df.columns if 'available' in col.lower()), None)
            if avail_column:
                df.loc[mask, avail_column] = availability
            
            # save file
            df.to_excel(file_path, index=False)
            
            # reload workers table
            self.load_workers_table(table)
            
            dialog.accept()
            
        except Exception as e:
            logging.error(f"Error updating worker: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error updating worker: {str(e)}")
    
    def delete_worker(self, table, email):
        """Delete worker from Excel file"""
        # confirm deletion
        confirm = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Are you sure you want to delete the worker with email {email}?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if confirm != QMessageBox.Yes:
            return
        
        try:
            # load Excel file
            file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Check if the worker exists
            if email not in df['Email'].values:
                QMessageBox.warning(self, "Warning", "Worker not found.")
                return
            
            # remove worker
            df = df[df['Email'] != email]
            
            # save file
            df.to_excel(file_path, index=False)
            
            # reload workers table
            self.load_workers_table(table)
            
            QMessageBox.information(self, "Success", "Worker deleted successfully.")
            
        except Exception as e:
            logging.error(f"Error deleting worker: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error deleting worker: {str(e)}")
    
    def manage_hours(self):
        """Show dialog to manage hours of operation"""
        # Get current hours of operation
        hours = {}
        if self.workplace in self.app_data and 'hours_of_operation' in self.app_data[self.workplace]:
            hours = self.app_data[self.workplace]['hours_of_operation']
        
        # Show dialog
        dialog = HoursOfOperationDialog(self.workplace, hours, self)
        
        if dialog.exec_() == QDialog.Accepted:
            # Save updated hours
            app_data = load_data()
            
            # Initialize workplace data if not exists
            if self.workplace not in app_data:
                app_data[self.workplace] = {}
            
            # Update hours of operation
            app_data[self.workplace]['hours_of_operation'] = dialog.hours_data
            
            # Save app data
            if save_data(app_data):
                # Update instance data
                self.app_data = app_data
                
                # Reload hours table
                self.load_hours_table(self.hours_table)
                
                QMessageBox.information(self, "Success", "Hours of operation saved successfully.")
            else:
                QMessageBox.critical(self, "Error", "Error saving hours of operation.")
    
    def generate_schedule(self):
        """Generate schedule for workplace"""
        # check if Excel file exists
        file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Warning", "No Excel file found for this workplace. Please upload one first.")
            return
        
        # check if hours of operation are defined
        if (self.workplace not in self.app_data or
            'hours_of_operation' not in self.app_data[self.workplace] or
            not any(blocks for blocks in self.app_data[self.workplace]['hours_of_operation'].values())):
            QMessageBox.warning(self, "Warning", "No hours of operation defined for this workplace. Please define them first.")
            return
        
        # show dialog to configure schedule generation
        dialog = QDialog(self)
        dialog.setWindowTitle("Generate Schedule")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # max hours per worker
        max_hours_spin = QSpinBox()
        max_hours_spin.setRange(1, 40)
        max_hours_spin.setValue(20)
        form_layout.addRow("Max Hours Per Worker:", max_hours_spin)
        
        # max workers per shift
        max_workers_spin = QSpinBox()
        max_workers_spin.setRange(1, 10)
        max_workers_spin.setValue(1)
        form_layout.addRow("Max Workers Per Shift:", max_workers_spin)
        
        layout.addLayout(form_layout)
        
        # buttons
        buttons_layout = QHBoxLayout()
        
        generate_btn = StyleHelper.create_button("Generate")
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        
        buttons_layout.addWidget(generate_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # connect buttons
        generate_btn.clicked.connect(lambda: self.do_generate_schedule(
            dialog,
            max_hours_spin.value(),
            max_workers_spin.value()
        ))
        
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec_()
    
    def do_generate_schedule(self, dialog, max_hours_per_worker, max_workers_per_shift):
        """Actually generate the schedule"""
        try:
            # load worker data
            file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Clean the DataFrame
            df = df.dropna(subset=['Email'], how='all')
            df = df[df['Email'].str.strip() != '']
            df = df[~df['Email'].str.contains('nan', case=False, na=False)]
            
            workers = []
            for _, row in df.iterrows():
                # Get availability from the "Days & Times Available" column
                avail_column = next((col for col in df.columns if 'available' in col.lower()), None)
                availability_text = str(row.get(avail_column, "")) if avail_column else ""
                if pd.isna(availability_text) or availability_text == "nan":
                    availability_text = ""
                
                # Parse availability into structured format
                availability = parse_availability(availability_text)
                
                workers.append({
                    "first_name": row.get("First Name", "").strip(),
                    "last_name": row.get("Last Name", "").strip(),
                    "email": row.get("Email", "").strip(),
                    "work_study": str(row.get("Work Study", "")).strip().lower() in ['yes', 'y', 'true'],
                    "availability": availability
                })
            
            # get hours of operation
            hours_of_operation = self.app_data[self.workplace]['hours_of_operation']
            
            # generate schedule
            schedule, assigned_hours, low_hour_workers, unassigned_workers, alternative_solutions, unfilled_shifts, work_study_issues = create_shifts_from_availability(
                hours_of_operation,
                workers,
                self.workplace,
                max_hours_per_worker,
                max_workers_per_shift
            )
            
            # close dialog
            dialog.accept()
            
            # Show alternative solutions dialog if there are unfilled shifts or work study issues
            if unfilled_shifts or work_study_issues:
                alt_dialog = AlternativeSolutionsDialog(alternative_solutions, unfilled_shifts, work_study_issues, self)
                alt_dialog.exec_()
            
            # show schedule
            self.show_schedule_dialog(schedule, assigned_hours, low_hour_workers, unassigned_workers, workers)
            
        except Exception as e:
            logging.error(f"Error generating schedule: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error generating schedule: {str(e)}")
    
    def show_schedule_dialog(self, schedule, assigned_hours, low_hour_workers, unassigned_workers, all_workers=None):
        """Show dialog with generated schedule"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Generated Schedule")
        dialog.setMinimumWidth(1000)
        dialog.setMinimumHeight(700)
        
        layout = QVBoxLayout()
        
        # schedule tabs
        tabs = QTabWidget()
        
        # schedule tab - using a single table for all days
        schedule_tab = QWidget()
        schedule_layout = QVBoxLayout(schedule_tab)
        
        # Create a single table for all shifts
        all_shifts_table = QTableWidget()
        all_shifts_table.setColumnCount(5)  # Day, Start, End, Assigned, Actions
        all_shifts_table.setHorizontalHeaderLabels(["Day", "Start", "End", "Assigned", "Actions"])
        
        # Count total shifts
        total_shifts = sum(len(shifts) for shifts in schedule.values())
        all_shifts_table.setRowCount(total_shifts)
        
        # Fill the table with all shifts
        row_index = 0
        for day in DAYS:
            shifts = schedule.get(day, [])
            for i, shift in enumerate(shifts):
                # Day
                day_item = QTableWidgetItem(day)
                all_shifts_table.setItem(row_index, 0, day_item)
                
                # Start
                start_item = QTableWidgetItem(format_time_ampm(shift['start']))
                all_shifts_table.setItem(row_index, 1, start_item)
                
                # End
                end_item = QTableWidgetItem(format_time_ampm(shift['end']))
                all_shifts_table.setItem(row_index, 2, end_item)
                
                # Assigned
                assigned_item = QTableWidgetItem(", ".join(shift['assigned']))
                if "Unfilled" in shift['assigned']:
                    assigned_item.setBackground(QColor(255, 200, 200))
                all_shifts_table.setItem(row_index, 3, assigned_item)
                
                # Edit button
                edit_widget = QWidget()
                edit_layout = QHBoxLayout(edit_widget)
                edit_layout.setContentsMargins(0, 0, 0, 0)
                
                edit_btn = QPushButton("Edit")
                edit_btn.setMinimumWidth(80)  # Make button wider
                edit_btn.setStyleSheet("background-color: #ffc107; color: black; font-size: 12px; padding: 6px 12px;")
                edit_btn.clicked.connect(lambda _, d=day, s=shift, r=row_index, t=all_shifts_table: 
                                        self.edit_shift_assignment(d, s, r, t, all_workers, dialog))
                
                edit_layout.addWidget(edit_btn)
                edit_layout.addStretch()
                all_shifts_table.setCellWidget(row_index, 4, edit_widget)
                
                row_index += 1
        
        # Set column widths
        all_shifts_table.setColumnWidth(0, 100)  # Day
        all_shifts_table.setColumnWidth(1, 100)  # Start
        all_shifts_table.setColumnWidth(2, 100)  # End
        all_shifts_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)  # Assigned
        all_shifts_table.setColumnWidth(4, 100)  # Actions
        
        schedule_layout.addWidget(all_shifts_table)
        
        # hours tab
        hours_tab = QWidget()
        hours_layout = QVBoxLayout(hours_tab)
        
        hours_table = QTableWidget()
        hours_table.setColumnCount(3)  # Added column for unassigned workers
        hours_table.setHorizontalHeaderLabels(["Worker", "Hours", "Status"])
        
        # Sort workers by hours (descending)
        sorted_workers = sorted(assigned_hours.items(), key=lambda x: x[1], reverse=True)
        
        # Include all workers, even those with 0 hours
        all_worker_emails = {w['email'] for w in all_workers}
        for email in all_worker_emails:
            if email not in assigned_hours:
                sorted_workers.append((email, 0))
        
        hours_table.setRowCount(len(sorted_workers))
        
        for i, (email, hours) in enumerate(sorted_workers):
            # Find worker name
            worker_name = email
            for worker in self.get_workers():
                if worker['email'] == email:
                    worker_name = f"{worker['first_name']} {worker['last_name']}"
                    break
            
            # Worker
            worker_item = QTableWidgetItem(worker_name)
            hours_table.setItem(i, 0, worker_item)
            
            # Hours
            hours_item = QTableWidgetItem(f"{hours:.1f}")
            if hours == 0:
                hours_item.setBackground(QColor(255, 200, 200))
            elif hours < 4:
                hours_item.setBackground(QColor(255, 255, 200))
            hours_table.setItem(i, 1, hours_item)
            
            # Status
            status = ""
            if hours == 0:
                status = "Unassigned"
                status_item = QTableWidgetItem(status)
                status_item.setBackground(QColor(255, 200, 200))
            elif hours < 4:
                status = "Low Hours"
                status_item = QTableWidgetItem(status)
                status_item.setBackground(QColor(255, 255, 200))
            else:
                status = "OK"
                status_item = QTableWidgetItem(status)
            
            hours_table.setItem(i, 2, status_item)
        
        hours_table.resizeColumnsToContents()
        hours_layout.addWidget(hours_table)
        
        # Low hour workers warning
        if low_hour_workers:
            warning_label = QLabel(f"Warning: The following workers have fewer than 4 hours: {', '.join(low_hour_workers)}")
            warning_label.setStyleSheet("color: red;")
            hours_layout.addWidget(warning_label)
        
        # Unassigned workers warning
        if unassigned_workers:
            unassigned_label = QLabel(f"Warning: The following workers have no assigned hours: {', '.join(unassigned_workers)}")
            unassigned_label.setStyleSheet("color: red; font-weight: bold;")
            hours_layout.addWidget(unassigned_label)
        
        # Add tabs
        tabs.addTab(schedule_tab, "Schedule")
        tabs.addTab(hours_tab, "Worker Hours")
        
        # Connect tab change signal to update worker hours
        tabs.currentChanged.connect(lambda: self.update_worker_hours_tab(dialog, hours_table))
        
        layout.addWidget(tabs)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        save_btn = StyleHelper.create_button("Save Schedule")
        email_btn = StyleHelper.create_button("Email Schedule")
        print_btn = StyleHelper.create_button("Print Schedule")
        close_btn = StyleHelper.create_button("Close", primary=False)
        
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(email_btn)
        buttons_layout.addWidget(print_btn)
        buttons_layout.addWidget(close_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # Store schedule and workers in dialog for editing
        dialog.schedule = schedule
        dialog.all_workers = all_workers
        dialog.assigned_hours = assigned_hours
        dialog.hours_table = hours_table
        
        # Connect buttons
        save_btn.clicked.connect(lambda: self.save_schedule(dialog, dialog.schedule))
        email_btn.clicked.connect(lambda: self.email_schedule_dialog(dialog.schedule))
        print_btn.clicked.connect(lambda: self.print_schedule(dialog.schedule))
        close_btn.clicked.connect(dialog.reject)
        
        dialog.exec_()
    
    def update_worker_hours_tab(self, dialog, hours_table):
        """Update the worker hours tab with current data"""
        if not hasattr(dialog, 'assigned_hours') or not hasattr(dialog, 'all_workers'):
            return
        
        # Recalculate assigned hours
        assigned_hours = {w['email']: 0 for w in dialog.all_workers}
        
        for day, shifts in dialog.schedule.items():
            for shift in shifts:
                shift_start = time_to_hour(shift['start'])
                shift_end = time_to_hour(shift['end'])
                shift_hours = shift_end - shift_start
                
                for email in shift.get('raw_assigned', []):
                    assigned_hours[email] = assigned_hours.get(email, 0) + shift_hours
        
        # Update the hours table
        sorted_workers = sorted(assigned_hours.items(), key=lambda x: x[1], reverse=True)
        
        for i, (email, hours) in enumerate(sorted_workers):
            if i < hours_table.rowCount():
                # Find worker name
                worker_name = email
                for worker in self.get_workers():
                    if worker['email'] == email:
                        worker_name = f"{worker['first_name']} {worker['last_name']}"
                        break
                
                # Update worker name
                hours_table.item(i, 0).setText(worker_name)
                
                # Update hours
                hours_table.item(i, 1).setText(f"{hours:.1f}")
                if hours == 0:
                    hours_table.item(i, 1).setBackground(QColor(255, 200, 200))
                elif hours < 4:
                    hours_table.item(i, 1).setBackground(QColor(255, 255, 200))
                else:
                    hours_table.item(i, 1).setBackground(QColor(255, 255, 255))
                
                # Update status
                if hours == 0:
                    status = "Unassigned"
                    hours_table.item(i, 2).setText(status)
                    hours_table.item(i, 2).setBackground(QColor(255, 200, 200))
                elif hours < 4:
                    status = "Low Hours"
                    hours_table.item(i, 2).setText(status)
                    hours_table.item(i, 2).setBackground(QColor(255, 255, 200))
                else:
                    status = "OK"
                    hours_table.item(i, 2).setText(status)
                    hours_table.item(i, 2).setBackground(QColor(255, 255, 255))
        
        # Update dialog's assigned hours
        dialog.assigned_hours = assigned_hours
    
    def edit_shift_assignment(self, day, shift, row, table, all_workers, parent_dialog):
        """Edit worker assignment for a shift"""
        if not all_workers:
            QMessageBox.warning(self, "Warning", "No workers available to edit this shift.")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Edit Shift: {day} {format_time_ampm(shift['start'])} - {format_time_ampm(shift['end'])}")
        dialog.setMinimumWidth(500)  # Increased width
        dialog.setMinimumHeight(500)  # Increased height
        
        layout = QVBoxLayout()
        
        # Get available workers for this shift
        available_workers = shift.get('all_available', [])
        
        # Add a label explaining what to do
        instruction_label = QLabel(f"Select workers for {day} {format_time_ampm(shift['start'])} - {format_time_ampm(shift['end'])}:")
        instruction_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(instruction_label)
        
        # Add a note about availability
        if not available_workers:
            no_workers_label = QLabel("No workers are available during this time slot based on their availability.")
            no_workers_label.setStyleSheet("color: red;")
            no_workers_label.setWordWrap(True)
            layout.addWidget(no_workers_label)
            
            # Show all workers anyway
            all_workers_label = QLabel("Showing all workers. You can assign them, but they may not be available during this time.")
            all_workers_label.setWordWrap(True)
            layout.addWidget(all_workers_label)
            
            # Use all workers instead
            available_workers = all_workers
        
        # Create list of workers
        worker_list = QListWidget()
        worker_list.setStyleSheet("QListWidget::item { padding: 5px; }")
        
        # Add all available workers
        for worker in available_workers:
            item = QListWidgetItem(f"{worker['first_name']} {worker['last_name']}")
            item.setData(Qt.UserRole, worker)
            
            # Check if worker is currently assigned
            if f"{worker['first_name']} {worker['last_name']}" in shift['assigned']:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            
            worker_list.setSelectionMode(QListWidget.NoSelection)
            worker_list.addItem(item)
        
        layout.addWidget(worker_list)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        save_btn = StyleHelper.create_button("Save")
        save_btn.setMinimumWidth(120)  # Wider button
        
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        cancel_btn.setMinimumWidth(120)  # Wider button
        
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
        save_btn.clicked.connect(lambda: self.update_shift_assignment(dialog, day, shift, row, table, worker_list, parent_dialog))
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec_()
    
    def update_shift_assignment(self, dialog, day, shift, row, table, worker_list, parent_dialog):
        """Update the worker assignment for a shift"""
        # Get selected workers
        selected_workers = []
        for i in range(worker_list.count()):
            item = worker_list.item(i)
            if item.checkState() == Qt.Checked:
                worker = item.data(Qt.UserRole)
                selected_workers.append(worker)
        
        # Update shift data
        shift['assigned'] = [f"{w['first_name']} {w['last_name']}" for w in selected_workers] if selected_workers else ["Unfilled"]
        shift['raw_assigned'] = [w['email'] for w in selected_workers] if selected_workers else []
        
        # Update table
        assigned_item = QTableWidgetItem(", ".join(shift['assigned']))
        if "Unfilled" in shift['assigned']:
            assigned_item.setBackground(QColor(255, 200, 200))
        table.setItem(row, 3, assigned_item)  # Update column 3 (Assigned) in the consolidated table
        
        # Update parent dialog's schedule
        if hasattr(parent_dialog, 'schedule'):
            parent_dialog.schedule[day] = [s if s != shift else shift for s in parent_dialog.schedule[day]]
            
            # Update worker hours tab if it's visible
            if hasattr(parent_dialog, 'hours_table'):
                self.update_worker_hours_tab(parent_dialog, parent_dialog.hours_table)
        
        dialog.accept()
    
    def get_workers(self):
        """Get workers from Excel file"""
        file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
        
        if not os.path.exists(file_path):
            return []
        
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Clean the DataFrame
            df = df.dropna(subset=['Email'], how='all')
            df = df[df['Email'].str.strip() != '']
            df = df[~df['Email'].str.contains('nan', case=False, na=False)]
            
            workers = []
            for _, row in df.iterrows():
                first_name = row.get("First Name", "").strip()
                last_name = row.get("Last Name", "").strip()
                email = row.get("Email", "").strip()
                work_study = str(row.get("Work Study", "")).strip().lower() in ['yes', 'y', 'true']
                
                # Skip invalid rows
                if pd.isna(email) or email == "" or email == "nan":
                    continue
                
                workers.append({
                    "first_name": first_name,
                    "last_name": last_name,
                    "email": email,
                    "work_study": work_study
                })
            
            return workers
        
        except Exception as e:
            logging.error(f"Error getting workers: {str(e)}")
            return []
    
    def save_schedule(self, dialog, schedule):
        """Save schedule to file"""
        try:
            # create save path for JSON
            json_path = os.path.join(DIRS['saved_schedules'], f"{self.workplace}_current.json")
            
            # save schedule as JSON
            with open(json_path, "w") as f:
                json.dump(schedule, f, indent=4)
            
            # Also save as Excel for easier reading
            excel_path = os.path.join(DIRS['saved_schedules'], f"{self.workplace}_current.xlsx")
            
            # Create Excel file
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                # Create a sheet for each day
                for day in DAYS:
                    if day in schedule and schedule[day]:
                        shifts = schedule[day]
                        data = []
                        for shift in shifts:
                            data.append({
                                "Start": format_time_ampm(shift['start']),
                                "End": format_time_ampm(shift['end']),
                                "Assigned": ", ".join(shift['assigned'])
                            })
                        df = pd.DataFrame(data)
                        df.to_excel(writer, sheet_name=day, index=False)
                
                # Create a summary sheet
                all_shifts = []
                for day, shifts in schedule.items():
                    for shift in shifts:
                        all_shifts.append({
                            "Day": day,
                            "Start": format_time_ampm(shift['start']),
                            "End": format_time_ampm(shift['end']),
                            "Assigned": ", ".join(shift['assigned'])
                        })
                
                if all_shifts:
                    summary_df = pd.DataFrame(all_shifts)
                    summary_df.to_excel(writer, sheet_name="Full Schedule", index=False)
            
            QMessageBox.information(dialog, "Success", f"Schedule saved successfully to:\n{excel_path}")
            
        except Exception as e:
            logging.error(f"Error saving schedule: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error saving schedule: {str(e)}")
    
    def email_schedule_dialog(self, schedule):
        """Show dialog to email schedule"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Email Schedule")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # sender email
        sender_email_input = QLineEdit()
        form_layout.addRow("Sender Email:", sender_email_input)
        
        # sender password
        sender_password_input = QLineEdit()
        sender_password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Sender Password:", sender_password_input)
        
        # Add note about Gmail app passwords
        gmail_note = QLabel("Note: For Gmail, you may need to use an App Password instead of your regular password. Go to your Google Account > Security > App Passwords to create one.")
        gmail_note.setWordWrap(True)
        gmail_note.setStyleSheet("font-style: italic; color: #666;")
        form_layout.addRow("", gmail_note)
        
        # recipients
        recipients_input = QTextEdit()
        recipients_input.setPlaceholderText("Enter email addresses, one per line")
        
        # add worker emails
        workers = self.get_workers()
        for worker in workers:
            if worker['email']:
                recipients_input.append(worker['email'])
        
        form_layout.addRow("Recipients:", recipients_input)
        
        layout.addLayout(form_layout)
        
        # buttons
        buttons_layout = QHBoxLayout()
        
        send_btn = StyleHelper.create_button("Send")
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        
        buttons_layout.addWidget(send_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # connect buttons
        send_btn.clicked.connect(lambda: self.send_schedule_email(
            dialog,
            schedule,
            sender_email_input.text(),
            sender_password_input.text(),
            recipients_input.toPlainText().splitlines()
        ))
        
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec_()
    
    def send_schedule_email(self, dialog, schedule, sender_email, sender_password, recipients):
        """Send schedule via email"""
        if not sender_email or not sender_password or not recipients:
            QMessageBox.warning(dialog, "Warning", "Sender email, password, and recipients are required.")
            return
        
        try:
            # send email
            success, message = send_schedule_email(
                self.workplace,
                schedule,
                recipients,
                sender_email,
                sender_password
            )
            
            if success:
                QMessageBox.information(dialog, "Success", message)
                dialog.accept()
            else:
                QMessageBox.critical(dialog, "Error", message)
            
        except Exception as e:
            logging.error(f"Error sending email: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error sending email: {str(e)}")
    
    def print_schedule(self, schedule):
        """Print the schedule"""
        try:
            # Create a printer
            printer = QPrinter()
            
            # Create a print dialog
            print_dialog = QPrintDialog(printer, self)
            if print_dialog.exec_() != QDialog.Accepted:
                return
            
            # Create the HTML content for printing
            html_content = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    h1 {{ text-align: center; }}
                    h2 {{ margin-top: 20px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
                    th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                    th {{ background-color: #f2f2f2; }}
                    .unfilled {{ color: red; }}
                </style>
            </head>
            <body>
                <h1>{self.workplace.replace('_', ' ').title()} Schedule</h1>
            """
            
            # Add each day's schedule
            for day in DAYS:
                if day in schedule and schedule[day]:
                    html_content += f"<h2>{day}</h2>"
                    html_content += "<table>"
                    html_content += "<tr><th>Start</th><th>End</th><th>Assigned</th></tr>"
                    
                    for shift in schedule[day]:
                        assigned = ", ".join(shift['assigned'])
                        unfilled_class = ' class="unfilled"' if "Unfilled" in assigned else ""
                        
                        html_content += "<tr>"
                        html_content += f"<td>{format_time_ampm(shift['start'])}</td>"
                        html_content += f"<td>{format_time_ampm(shift['end'])}</td>"
                        html_content += f"<td{unfilled_class}>{assigned}</td>"
                        html_content += "</tr>"
                    
                    html_content += "</table>"
            
            html_content += """
            </body>
            </html>
            """
            
            # Create a QTextDocument to render the HTML
            from PyQt5.QtGui import QTextDocument
            document = QTextDocument()
            document.setHtml(html_content)
            
            # Print the document
            document.print_(printer)
            
            QMessageBox.information(self, "Success", "Schedule sent to printer.")
            
        except Exception as e:
            logging.error(f"Error printing schedule: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error printing schedule: {str(e)}")
    
    def view_current_schedule(self):
        """View current saved schedule"""
        # check if schedule exists
        save_path = os.path.join(DIRS['saved_schedules'], f"{self.workplace}_current.json")
        
        if not os.path.exists(save_path):
            QMessageBox.warning(self, "Warning", "No saved schedule found for this workplace.")
            return
        
        try:
            # load schedule
            with open(save_path, "r") as f:
                schedule = json.load(f)
            
            # load workers for editing
            all_workers = self.get_workers()
            
            # calculate assigned hours
            assigned_hours = {}
            for day, shifts in schedule.items():
                for shift in shifts:
                    shift_start = time_to_hour(shift['start'])
                    shift_end = time_to_hour(shift['end'])
                    shift_hours = shift_end - shift_start
                    
                    for email in shift.get('raw_assigned', []):
                        assigned_hours[email] = assigned_hours.get(email, 0) + shift_hours
            
            # identify unassigned workers
            unassigned_workers = []
            for w in all_workers:
                if assigned_hours.get(w['email'], 0) == 0:
                    unassigned_workers.append(f"{w['first_name']} {w['last_name']}")
            
            # identify low hour workers
            low_hour_workers = []
            for w in all_workers:
                if assigned_hours.get(w['email'], 0) > 0 and assigned_hours.get(w['email'], 0) < 4:
                    low_hour_workers.append(f"{w['first_name']} {w['last_name']}")
            
            # show schedule
            self.show_schedule_dialog(schedule, assigned_hours, low_hour_workers, unassigned_workers, all_workers)
            
        except Exception as e:
            logging.error(f"Error viewing schedule: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error viewing schedule: {str(e)}")
    
    def show_last_minute_dialog(self):
        """Show dialog for last minute availability"""
        dialog = LastMinuteAvailabilityDialog(self.workplace, self)
        dialog.exec_()
    
    def check_last_minute_availability(self):
        """Check last minute availability from the tab"""
        day = self.lm_day_combo.currentText()
        start_time = self.lm_start_time.time().toString("HH:mm")
        end_time = self.lm_end_time.time().toString("HH:mm")
        
        # Get workers
        workers = self.get_workers()
        
        # Find available workers
        available_workers = []
        for worker in workers:
            # Get availability
            avail_column = None
            file_path = os.path.join(DIRS['workplaces'], f"{self.workplace}.xlsx")
            if os.path.exists(file_path):
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                avail_column = next((col for col in df.columns if 'available' in col.lower()), None)
                
                if avail_column:
                    worker_row = df[df['Email'] == worker['email']]
                    if not worker_row.empty:
                        availability_text = str(worker_row.iloc[0].get(avail_column, ""))
                        if not pd.isna(availability_text) and availability_text != "nan":
                            availability = parse_availability(availability_text)
                            
                            # Check if worker is available
                            start_hour = time_to_hour(start_time)
                            end_hour = time_to_hour(end_time)
                            
                            day_availability = availability.get(day, [])
                            for avail in day_availability:
                                avail_start = avail['start_hour']
                                avail_end = avail['end_hour']
                                
                                if avail_start <= start_hour and end_hour <= avail_end:
                                    available_workers.append(worker)
                                    break
        
        # Display results
        self.lm_results_table.setRowCount(len(available_workers))
        
        for i, worker in enumerate(available_workers):
            # Name
            name = f"{worker['first_name']} {worker['last_name']}"
            self.lm_results_table.setItem(i, 0, QTableWidgetItem(name))
            
            # Email
            self.lm_results_table.setItem(i, 1, QTableWidgetItem(worker['email']))
            
            # Work Study
            work_study = "Yes" if worker['work_study'] else "No"
            self.lm_results_table.setItem(i, 2, QTableWidgetItem(work_study))
        
        # Show message if no workers are available
        if not available_workers:
            QMessageBox.warning(self, "No Available Workers", 
                               f"No workers are available on {day} from {format_time_ampm(start_time)} to {format_time_ampm(end_time)}.")

class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(1000, 700)
        
        # set style
        self.setStyleSheet(StyleHelper.get_main_style())
        
        # central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # main layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # header
        header_layout = QHBoxLayout()
        
        title_label = QLabel(APP_NAME)
        font = title_label.font()
        font.setPointSize(16)
        font.setBold(True)
        title_label.setFont(font)
        
        version_label = QLabel(f"v{APP_VERSION}")
        
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(version_label)
        
        main_layout.addLayout(header_layout)
        
        # tabs
        tabs = QTabWidget()
        
        # add workplace tabs
        tabs.addTab(WorkplaceTab("esports_lounge"), "eSports Lounge")
        tabs.addTab(WorkplaceTab("esports_arena"), "eSports Arena")
        tabs.addTab(WorkplaceTab("it_service_center"), "IT Service Center")
        
        main_layout.addWidget(tabs)
        
        # show window
        self.show()

# main function
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
