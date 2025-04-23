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
matplotlib.use('Agg')  # Use non-interactive backend
from datetime import datetime, time, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QPushButton, QComboBox, QTableWidget, QTableWidgetItem,
                            QFileDialog, QMessageBox, QTabWidget, QLineEdit, QCheckBox,
                            QTimeEdit, QSpinBox, QFormLayout, QGroupBox, QTextEdit, QDialog,
                            QScrollArea, QFrame, QSplitter, QStackedWidget, QListWidget)
from PyQt5.QtCore import Qt, QTime, QSize, QSettings, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QIcon, QPixmap, QColor, QPalette

# Constants
DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
DATA_FILE = 'data.json'
APP_NAME = "Workplace Scheduler"
APP_VERSION = "1.0.0"

# Setup logging
logging.basicConfig(
    filename='logs/app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Ensure directories exist
os.makedirs('workplaces', exist_ok=True)
os.makedirs('schedules', exist_ok=True)
os.makedirs('saved_schedules', exist_ok=True)
os.makedirs('logs', exist_ok=True)

# Utility Functions
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

def parse_not_available(raw_string):
    """Parse not available times from string format"""
    if pd.isna(raw_string) or not raw_string:
        return {}
        
    day_map = {
        "sun": "Sunday", "mon": "Monday", "tue": "Tuesday",
        "wed": "Wednesday", "thu": "Thursday", "fri": "Friday", "sat": "Saturday"
    }
    not_available = {}
    
    blocks = re.split(r',\s*', str(raw_string))
    for block in blocks:
        match = re.match(r'(\w+)\s*(\d{1,2})\s*[-to]+\s*(\d{1,2})', block, re.IGNORECASE)
        if match:
            day_raw, start, end = match.groups()
            day_key = day_map.get(day_raw[:3].lower(), None)
            if day_key:
                not_available.setdefault(day_key, []).append((int(start), int(end)))
    
    return not_available

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
    """Check if a worker is available for a given shift"""
    # First check if they're generally available that day
    if not worker['availability'].get(day, False):
        return False
    
    # Then check if the shift overlaps with any of their unavailable times
    for block in worker.get('not_available', {}).get(day, []):
        if overlaps(shift_start, shift_end, block[0], block[1]):
            return False
    
    return True

def create_shifts_from_availability(hours_of_operation, workers, workplace, max_hours_per_worker, max_workers_per_shift):
    """Create shifts based on hours of operation and worker availability"""
    schedule = {}
    
    # Determine ideal shift length based on workplace
    if "lounge" in workplace.lower() or "arena" in workplace.lower():
        ideal_shift_length = 3  # 3-hour shifts for lounges/arenas
    else:
        ideal_shift_length = 4  # 4-hour shifts for other workplaces
    
    min_shift_length = 2  # Minimum shift length in hours
    
    # Track assigned hours per worker
    assigned_hours = {w['email']: 0 for w in workers}
    assigned_days = {w['email']: set() for w in workers}
    
    # Track if a worker is work study (limited to 5 hours per week)
    work_study_status = {w['email']: w['work_study'] for w in workers}
    
    # For each day in the week
    for day, operation_hours in hours_of_operation.items():
        if not operation_hours:
            continue  # Skip days with no hours of operation
            
        schedule[day] = []
        
        # For each operation period in the day (e.g., morning and evening blocks)
        for op in operation_hours:
            start_hour = time_to_hour(op['start'])
            end_hour = time_to_hour(op['end'])
            
            # Skip if invalid hours
            if end_hour <= start_hour:
                end_hour += 24  # Handle overnight shifts
            
            total_hours = end_hour - start_hour
            
            # Create shifts based on ideal shift length
            current_hour = start_hour
            while current_hour < end_hour:
                # Determine shift end time (not exceeding operation end)
                shift_end = min(current_hour + ideal_shift_length, end_hour)
                
                # If remaining time is too short, extend the previous shift
                if end_hour - shift_end < min_shift_length and shift_end < end_hour:
                    shift_end = end_hour
                
                # Only create the shift if it meets minimum length
                if shift_end - current_hour >= min_shift_length:
                    # Find available workers for this shift
                    available_workers = []
                    for worker in workers:
                        email = worker['email']
                        
                        # Check if worker is available
                        if is_worker_available(worker, day, current_hour, shift_end):
                            # Check work study limit (5 hours per week)
                            if work_study_status[email] and assigned_hours.get(email, 0) + (shift_end - current_hour) > 5:
                                continue
                                
                            # Check max hours per worker limit
                            if assigned_hours.get(email, 0) + (shift_end - current_hour) > max_hours_per_worker:
                                continue
                                
                            # Add to available workers
                            available_workers.append(worker)
                    
                    # Sort workers by assigned hours (least to most)
                    available_workers.sort(key=lambda w: (assigned_hours[w['email']], random.random()))
                    
                    # Assign workers to shift (up to max_workers_per_shift)
                    assigned = []
                    for worker in available_workers[:max_workers_per_shift]:
                        assigned.append(worker)
                        
                        # Update worker's hours
                        email = worker['email']
                        assigned_hours[email] += (shift_end - current_hour)
                        assigned_days[email].add(day)
                    
                    # Add shift to schedule
                    schedule[day].append({
                        "start": hour_to_time_str(current_hour),
                        "end": hour_to_time_str(shift_end),
                        "assigned": [f"{w['first_name']} {w['last_name']}" for w in assigned] if assigned else ["Unfilled"],
                        "available": [f"{w['first_name']} {w['last_name']}" for w in available_workers],
                        "raw_assigned": [w['email'] for w in assigned] if assigned else []
                    })
                
                # Move to next shift
                current_hour = shift_end
    
    # Identify workers with low hours
    low_hour_workers = []
    for w in workers:
        if not w['work_study'] and assigned_hours[w['email']] < 4:
            low_hour_workers.append(f"{w['first_name']} {w['last_name']}")
    
    return schedule, assigned_hours, low_hour_workers

def send_schedule_email(workplace, schedule, recipient_emails, sender_email, sender_password):
    """Send schedule via email"""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ", ".join(recipient_emails)
        msg['Subject'] = f"{workplace.replace('_', ' ').title()} Schedule"
        
        # Create HTML body
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
        
        # Add schedule tables by day
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
        
        # Attach HTML body
        msg.attach(MIMEText(html, 'html'))
        
        # Create schedule image
        img_path = create_schedule_image(workplace, schedule)
        if img_path and os.path.exists(img_path):
            with open(img_path, 'rb') as f:
                img = MIMEImage(f.read())
                img.add_header('Content-Disposition', 'attachment', filename=f"{workplace}_schedule.png")
                msg.attach(img)
        
        # Create CSV file
        csv_path = create_schedule_csv(workplace, schedule)
        if csv_path and os.path.exists(csv_path):
            with open(csv_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype="csv")
                attachment.add_header('Content-Disposition', 'attachment', filename=f"{workplace}_schedule.csv")
                msg.attach(attachment)
        
        # Send email
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        
        return True, "Email sent successfully"
    
    except Exception as e:
        logging.error(f"Error sending email: {str(e)}")
        return False, f"Error sending email: {str(e)}"

def create_schedule_image(workplace, schedule):
    """Create an image of the schedule"""
    try:
        # Flatten schedule into rows
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
        
        # Create figure and axis
        fig, ax = plt.subplots(figsize=(10, len(rows) * 0.4))
        ax.axis('off')
        
        # Create table data
        table_data = [["Day", "Start", "End", "Assigned"]] + [[r["Day"], format_time_ampm(r["Start"]), format_time_ampm(r["End"]), r["Assigned"]] for r in rows]
        
        # Create table
        table = ax.table(cellText=table_data, cellLoc='center', loc='center')
        
        # Style table
        for cell in table.get_celld().values():
            cell.set_fontsize(10)
        
        # Save figure
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join("schedules", f"{workplace}_{timestamp}.png")
        plt.savefig(output_path, bbox_inches='tight')
        plt.close()
        
        return output_path
    
    except Exception as e:
        logging.error(f"Error creating schedule image: {str(e)}")
        return None

def create_schedule_csv(workplace, schedule):
    """Create a CSV file of the schedule"""
    try:
        # Flatten schedule into rows
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
        
        # Create DataFrame
        df = pd.DataFrame(rows)
        
        # Save to CSV
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join("schedules", f"{workplace}_{timestamp}.csv")
        df.to_csv(output_path, index=False)
        
        return output_path
    
    except Exception as e:
        logging.error(f"Error creating schedule CSV: {str(e)}")
        return None

# Main Application Classes
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
            QLineEdit, QComboBox, QSpinBox, QTimeEdit {
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

class WorkplaceTab(QWidget):
    """Tab for managing a specific workplace"""
    
    def __init__(self, workplace, parent=None):
        super().__init__(parent)
        self.workplace = workplace
        self.app_data = load_data()
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout()
        
        # Workplace title
        title = StyleHelper.create_section_title(f"{self.workplace.replace('_', ' ').title()}")
        layout.addWidget(title)
        
        # Quick actions
        actions_layout = QHBoxLayout()
        
        upload_btn = StyleHelper.create_button("Upload Excel File")
        upload_btn.clicked.connect(self.upload_excel)
        
        hours_btn = StyleHelper.create_button("Hours of Operation")
        hours_btn.clicked.connect(self.manage_hours)
        
        generate_btn = StyleHelper.create_action_button("Generate Schedule")
        generate_btn.clicked.connect(self.generate_schedule)
        
        view_btn = StyleHelper.create_button("View Current Schedule", primary=False)
        view_btn.clicked.connect(self.view_current_schedule)
        
        actions_layout.addWidget(upload_btn)
        actions_layout.addWidget(hours_btn)
        actions_layout.addWidget(generate_btn)
        actions_layout.addWidget(view_btn)
        
        layout.addLayout(actions_layout)
        
        # Tab widget for different sections
        tabs = QTabWidget()
        
        # Workers tab
        workers_tab = QWidget()
        workers_layout = QVBoxLayout()
        
        # Workers table
        workers_table = QTableWidget()
        workers_table.setColumnCount(6)
        workers_table.setHorizontalHeaderLabels(["First Name", "Last Name", "Email", "Work Study", "Availability", "Actions"])
        
        # Load workers
        self.load_workers_table(workers_table)
        
        workers_layout.addWidget(workers_table)
        
        # Add worker button
        add_worker_btn = StyleHelper.create_button("Add Worker")
        add_worker_btn.clicked.connect(lambda: self.add_worker_dialog(workers_table))
        workers_layout.addWidget(add_worker_btn)
        
        workers_tab.setLayout(workers_layout)
        
        # Hours of Operation tab
        hours_tab = QWidget()
        hours_layout = QVBoxLayout()
        
        hours_group = QGroupBox("Current Hours of Operation")
        hours_group_layout = QVBoxLayout()
        
        # Load hours of operation
        hours_table = QTableWidget()
        hours_table.setColumnCount(3)
        hours_table.setHorizontalHeaderLabels(["Day", "Start", "End"])
        
        self.load_hours_table(hours_table)
        
        hours_group_layout.addWidget(hours_table)
        hours_group.setLayout(hours_group_layout)
        hours_layout.addWidget(hours_group)
        
        edit_hours_btn = StyleHelper.create_button("Edit Hours of Operation")
        edit_hours_btn.clicked.connect(self.manage_hours)
        hours_layout.addWidget(edit_hours_btn)
        
        hours_tab.setLayout(hours_layout)
        
        # Add tabs
        tabs.addTab(workers_tab, "Workers")
        tabs.addTab(hours_tab, "Hours of Operation")
        
        layout.addWidget(tabs)
        
        self.setLayout(layout)
    
    def load_workers_table(self, table):
        """Load workers into table"""
        # Clear table
        table.setRowCount(0)
        
        # Check if Excel file exists
        file_path = os.path.join("workplaces", f"{self.workplace}.xlsx")
        if not os.path.exists(file_path):
            return
        
        # Load Excel file
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Set row count
            table.setRowCount(len(df))
            
            # Fill table
            for i, (_, row) in enumerate(df.iterrows()):
                # First Name
                first_name = row.get("First Name", "")
                table.setItem(i, 0, QTableWidgetItem(str(first_name)))
                
                # Last Name
                last_name = row.get("Last Name", "")
                table.setItem(i, 1, QTableWidgetItem(str(last_name)))
                
                # Email
                email = row.get("Email", "")
                table.setItem(i, 2, QTableWidgetItem(str(email)))
                
                # Work Study
                work_study = row.get("Work Study", "No")
                table.setItem(i, 3, QTableWidgetItem(str(work_study)))
                
                # Availability
                availability = []
                for day in DAYS:
                    if str(row.get(day, "No")).strip().lower() in ['yes', 'y', 'true']:
                        availability.append(day[:3])
                
                table.setItem(i, 4, QTableWidgetItem(", ".join(availability)))
                
                # Actions
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
            
            # Resize columns
            table.resizeColumnsToContents()
            
        except Exception as e:
            logging.error(f"Error loading workers: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error loading workers: {str(e)}")
    
    def load_hours_table(self, table):
        """Load hours of operation into table"""
        # Clear table
        table.setRowCount(0)
        
        # Get hours of operation
        hours = {}
        if self.workplace in self.app_data and 'hours_of_operation' in self.app_data[self.workplace]:
            hours = self.app_data[self.workplace]['hours_of_operation']
        
        # Count total rows needed
        total_rows = sum(len(blocks) if blocks else 1 for blocks in hours.values())
        table.setRowCount(total_rows)
        
        # Fill table
        row_index = 0
        for day in DAYS:
            blocks = hours.get(day, [])
            
            if not blocks:
                # No hours for this day
                table.setItem(row_index, 0, QTableWidgetItem(day))
                table.setItem(row_index, 1, QTableWidgetItem("Closed"))
                table.setItem(row_index, 2, QTableWidgetItem("Closed"))
                row_index += 1
            else:
                # Hours for this day
                for block in blocks:
                    table.setItem(row_index, 0, QTableWidgetItem(day))
                    table.setItem(row_index, 1, QTableWidgetItem(format_time_ampm(block['start'])))
                    table.setItem(row_index, 2, QTableWidgetItem(format_time_ampm(block['end'])))
                    row_index += 1
        
        # Resize columns
        table.resizeColumnsToContents()
    
    def upload_excel(self):
        """Upload Excel file for workplace"""
        file_path, _ = QFileDialog.getOpenFileName(self, "Upload Excel File", "", "Excel Files (*.xlsx)")
        
        if not file_path:
            return
        
        try:
            # Copy file to workplaces directory
            destination = os.path.join("workplaces", f"{self.workplace}.xlsx")
            import shutil
            shutil.copy2(file_path, destination)
            
            # Reload workers table
            workers_table = self.findChild(QTableWidget)
            if workers_table:
                self.load_workers_table(workers_table)
            
            QMessageBox.information(self, "Success", "Excel file uploaded successfully.")
            
        except Exception as e:
            logging.error(f"Error uploading Excel file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error uploading Excel file: {str(e)}")
    
    def add_worker_dialog(self, table):
        """Show dialog to add a worker"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Worker")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # First Name
        first_name_input = QLineEdit()
        form_layout.addRow("First Name:", first_name_input)
        
        # Last Name
        last_name_input = QLineEdit()
        form_layout.addRow("Last Name:", last_name_input)
        
        # Email
        email_input = QLineEdit()
        form_layout.addRow("Email:", email_input)
        
        # Work Study
        work_study_combo = QComboBox()
        work_study_combo.addItems(["No", "Yes"])
        form_layout.addRow("Work Study:", work_study_combo)
        
        # Availability
        avail_group = QGroupBox("Availability")
        avail_layout = QVBoxLayout()
        
        day_checkboxes = {}
        for day in DAYS:
            checkbox = QCheckBox(day)
            avail_layout.addWidget(checkbox)
            day_checkboxes[day] = checkbox
        
        avail_group.setLayout(avail_layout)
        form_layout.addRow("", avail_group)
        
        # Not Available
        not_avail_input = QLineEdit()
        not_avail_input.setPlaceholderText("e.g. Mon 12-2, Wed 1-3")
        form_layout.addRow("Not Available:", not_avail_input)
        
        layout.addLayout(form_layout)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        save_btn = StyleHelper.create_button("Save")
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
        save_btn.clicked.connect(lambda: self.save_worker(
            dialog,
            table,
            first_name_input.text(),
            last_name_input.text(),
            email_input.text(),
            work_study_combo.currentText(),
            {day: checkbox.isChecked() for day, checkbox in day_checkboxes.items()},
            not_avail_input.text()
        ))
        
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec_()
    
    def save_worker(self, dialog, table, first_name, last_name, email, work_study, availability, not_available):
        """Save worker to Excel file"""
        if not first_name or not last_name or not email:
            QMessageBox.warning(dialog, "Warning", "First name, last name, and email are required.")
            return
        
        try:
            # Check if Excel file exists
            file_path = os.path.join("workplaces", f"{self.workplace}.xlsx")
            
            if os.path.exists(file_path):
                # Load existing file
                df = pd.read_excel(file_path)
                df.columns = df.columns.str.strip()
                
                # Check if email already exists
                if email in df['Email'].values:
                    QMessageBox.warning(dialog, "Warning", "A worker with this email already exists.")
                    return
                
                # Create new row
                new_row = {
                    "First Name": first_name,
                    "Last Name": last_name,
                    "Email": email,
                    "Work Study": work_study
                }
                
                # Add availability
                for day in DAYS:
                    new_row[day] = "Yes" if availability.get(day, False) else "No"
                
                # Add not available
                not_avail_column = next((col for col in df.columns if 'not avail' in col.lower()), None)
                if not_avail_column:
                    new_row[not_avail_column] = not_available
                
                # Append row
                df = df.append(new_row, ignore_index=True)
                
            else:
                # Create new file
                columns = ["First Name", "Last Name", "Email", "Work Study"] + DAYS + ["Days / Times Not Available"]
                df = pd.DataFrame(columns=columns)
                
                # Create new row
                new_row = {
                    "First Name": first_name,
                    "Last Name": last_name,
                    "Email": email,
                    "Work Study": work_study
                }
                
                # Add availability
                for day in DAYS:
                    new_row[day] = "Yes" if availability.get(day, False) else "No"
                
                # Add not available
                new_row["Days / Times Not Available"] = not_available
                
                # Append row
                df = df.append(new_row, ignore_index=True)
            
            # Save file
            df.to_excel(file_path, index=False)
            
            # Reload workers table
            self.load_workers_table(table)
            
            dialog.accept()
            
        except Exception as e:
            logging.error(f"Error saving worker: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error saving worker: {str(e)}")
    
    def edit_worker_dialog(self, table, row, email):
        """Show dialog to edit a worker"""
        # Get worker data
        file_path = os.path.join("workplaces", f"{self.workplace}.xlsx")
        
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Warning", "Excel file not found.")
            return
        
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Find worker
            worker_row = df[df['Email'] == email]
            
            if worker_row.empty:
                QMessageBox.warning(self, "Warning", "Worker not found.")
                return
            
            worker_row = worker_row.iloc[0]
            
            # Create dialog
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Worker")
            dialog.setMinimumWidth(400)
            
            layout = QVBoxLayout()
            
            form_layout = QFormLayout()
            
            # First Name
            first_name_input = QLineEdit()
            first_name_input.setText(str(worker_row.get("First Name", "")))
            form_layout.addRow("First Name:", first_name_input)
            
            # Last Name
            last_name_input = QLineEdit()
            last_name_input.setText(str(worker_row.get("Last Name", "")))
            form_layout.addRow("Last Name:", last_name_input)
            
            # Email
            email_input = QLineEdit()
            email_input.setText(str(worker_row.get("Email", "")))
            email_input.setReadOnly(True)  # Email cannot be changed
            form_layout.addRow("Email:", email_input)
            
            # Work Study
            work_study_combo = QComboBox()
            work_study_combo.addItems(["No", "Yes"])
            work_study_combo.setCurrentText(str(worker_row.get("Work Study", "No")))
            form_layout.addRow("Work Study:", work_study_combo)
            
            # Availability
            avail_group = QGroupBox("Availability")
            avail_layout = QVBoxLayout()
            
            day_checkboxes = {}
            for day in DAYS:
                checkbox = QCheckBox(day)
                checkbox.setChecked(str(worker_row.get(day, "No")).strip().lower() in ['yes', 'y', 'true'])
                avail_layout.addWidget(checkbox)
                day_checkboxes[day] = checkbox
            
            avail_group.setLayout(avail_layout)
            form_layout.addRow("", avail_group)
            
            # Not Available
            not_avail_column = next((col for col in df.columns if 'not avail' in col.lower()), None)
            not_avail_input = QLineEdit()
            if not_avail_column:
                not_avail_input.setText(str(worker_row.get(not_avail_column, "")))
            not_avail_input.setPlaceholderText("e.g. Mon 12-2, Wed 1-3")
            form_layout.addRow("Not Available:", not_avail_input)
            
            layout.addLayout(form_layout)
            
            # Buttons
            buttons_layout = QHBoxLayout()
            
            save_btn = StyleHelper.create_button("Save")
            cancel_btn = StyleHelper.create_button("Cancel", primary=False)
            
            buttons_layout.addWidget(save_btn)
            buttons_layout.addWidget(cancel_btn)
            
            layout.addLayout(buttons_layout)
            
            dialog.setLayout(layout)
            
            # Connect buttons
            save_btn.clicked.connect(lambda: self.update_worker(
                dialog,
                table,
                email,
                first_name_input.text(),
                last_name_input.text(),
                work_study_combo.currentText(),
                {day: checkbox.isChecked() for day, checkbox in day_checkboxes.items()},
                not_avail_input.text()
            ))
            
            cancel_btn.clicked.connect(dialog.reject)
            
            dialog.exec_()
            
        except Exception as e:
            logging.error(f"Error editing worker: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error editing worker: {str(e)}")
    
    def update_worker(self, dialog, table, email, first_name, last_name, work_study, availability, not_available):
        """Update worker in Excel file"""
        if not first_name or not last_name:
            QMessageBox.warning(dialog, "Warning", "First name and last name are required.")
            return
        
        try:
            # Load Excel file
            file_path = os.path.join("workplaces", f"{self.workplace}.xlsx")
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Find worker
            mask = df['Email'] == email
            
            if not any(mask):
                QMessageBox.warning(dialog, "Warning", "Worker not found.")
                return
            
            # Update worker
            df.loc[mask, "First Name"] = first_name
            df.loc[mask, "Last Name"] = last_name
            df.loc[mask, "Work Study"] = work_study
            
            # Update availability
            for day in DAYS:
                df.loc[mask, day] = "Yes" if availability.get(day, False) else "No"
            
            # Update not available
            not_avail_column = next((col for col in df.columns if 'not avail' in col.lower()), None)
            if not_avail_column:
                df.loc[mask, not_avail_column] = not_available
            
            # Save file
            df.to_excel(file_path, index=False)
            
            # Reload workers table
            self.load_workers_table(table)
            
            dialog.accept()
            
        except Exception as e:
            logging.error(f"Error updating worker: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error updating worker: {str(e)}")
    
    def delete_worker(self, table, email):
        """Delete worker from Excel file"""
        # Confirm deletion
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
            # Load Excel file
            file_path = os.path.join("workplaces", f"{self.workplace}.xlsx")
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Remove worker
            df = df[df['Email'] != email]
            
            # Save file
            df.to_excel(file_path, index=False)
            
            # Reload workers table
            self.load_workers_table(table)
            
        except Exception as e:
            logging.error(f"Error deleting worker: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error deleting worker: {str(e)}")
    
    def manage_hours(self):
        """Show dialog to manage hours of operation"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Hours of Operation")
        dialog.setMinimumWidth(500)
        
        layout = QVBoxLayout()
        
        # Get current hours of operation
        hours = {}
        if self.workplace in self.app_data and 'hours_of_operation' in self.app_data[self.workplace]:
            hours = self.app_data[self.workplace]['hours_of_operation']
        
        # Create form
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout()
        
        day_widgets = {}
        
        for day in DAYS:
            day_group = QGroupBox(day)
            day_layout = QVBoxLayout()
            
            # Get blocks for this day
            blocks = hours.get(day, [])
            
            # Create widgets for blocks
            block_widgets = []
            
            if blocks:
                for block in blocks:
                    block_layout = QHBoxLayout()
                    
                    start_time = QTimeEdit()
                    start_time.setDisplayFormat("HH:mm")
                    if 'start' in block:
                        start_time.setTime(QTime.fromString(block['start'], "HH:mm"))
                    
                    end_time = QTimeEdit()
                    end_time.setDisplayFormat("HH:mm")
                    if 'end' in block:
                        end_time.setTime(QTime.fromString(block['end'], "HH:mm"))
                    
                    block_layout.addWidget(QLabel("Start:"))
                    block_layout.addWidget(start_time)
                    block_layout.addWidget(QLabel("End:"))
                    block_layout.addWidget(end_time)
                    
                    remove_btn = QPushButton("Remove")
                    remove_btn.setStyleSheet("background-color: #dc3545;")
                    remove_btn.clicked.connect(lambda _, layout=block_layout: layout.parentWidget().deleteLater())
                    
                    block_layout.addWidget(remove_btn)
                    
                    block_widget = QWidget()
                    block_widget.setLayout(block_layout)
                    day_layout.addWidget(block_widget)
                    
                    block_widgets.append((start_time, end_time))
            
            # Add button
            add_btn = QPushButton("Add Time Block")
            add_btn.clicked.connect(lambda _, d=day, l=day_layout, w=block_widgets: self.add_time_block(d, l, w))
            
            day_layout.addWidget(add_btn)
            day_group.setLayout(day_layout)
            
            scroll_layout.addWidget(day_group)
            
            day_widgets[day] = block_widgets
        
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        
        layout.addWidget(scroll_area)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        save_btn = StyleHelper.create_button("Save")
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
        save_btn.clicked.connect(lambda: self.save_hours(dialog, day_widgets))
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec_()
    
    def add_time_block(self, day, layout, block_widgets):
        """Add a time block to a day"""
        block_layout = QHBoxLayout()
        
        start_time = QTimeEdit()
        start_time.setDisplayFormat("HH:mm")
        
        end_time = QTimeEdit()
        end_time.setDisplayFormat("HH:mm")
        
        block_layout.addWidget(QLabel("Start:"))
        block_layout.addWidget(start_time)
        block_layout.addWidget(QLabel("End:"))
        block_layout.addWidget(end_time)
        
        remove_btn = QPushButton("Remove")
        remove_btn.setStyleSheet("background-color: #dc3545;")
        remove_btn.clicked.connect(lambda _, layout=block_layout: layout.parentWidget().deleteLater())
        
        block_layout.addWidget(remove_btn)
        
        block_widget = QWidget()
        block_widget.setLayout(block_layout)
        
        # Insert before the Add button
        layout.insertWidget(layout.count() - 1, block_widget)
        
        block_widgets.append((start_time, end_time))
    
    def save_hours(self, dialog, day_widgets):
        """Save hours of operation"""
        try:
            # Load app data
            app_data = load_data()
            
            # Initialize workplace data if not exists
            if self.workplace not in app_data:
                app_data[self.workplace] = {}
            
            # Initialize hours of operation if not exists
            if 'hours_of_operation' not in app_data[self.workplace]:
                app_data[self.workplace]['hours_of_operation'] = {}
            
            # Update hours of operation
            for day, block_widgets in day_widgets.items():
                blocks = []
                
                for start_time, end_time in block_widgets:
                    start = start_time.time().toString("HH:mm")
                    end = end_time.time().toString("HH:mm")
                    
                    blocks.append({
                        "start": start,
                        "end": end
                    })
                
                app_data[self.workplace]['hours_of_operation'][day] = blocks
            
            # Save app data
            if save_data(app_data):
                # Update instance data
                self.app_data = app_data
                
                # Reload hours table
                hours_table = self.findChild(QTableWidget, "", Qt.FindChildrenRecursively)
                if hours_table:
                    self.load_hours_table(hours_table)
                
                dialog.accept()
                
                QMessageBox.information(self, "Success", "Hours of operation saved successfully.")
            else:
                QMessageBox.critical(dialog, "Error", "Error saving hours of operation.")
            
        except Exception as e:
            logging.error(f"Error saving hours of operation: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error saving hours of operation: {str(e)}")
    
    def generate_schedule(self):
        """Generate schedule for workplace"""
        # Check if Excel file exists
        file_path = os.path.join("workplaces", f"{self.workplace}.xlsx")
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Warning", "No Excel file found for this workplace. Please upload one first.")
            return
        
        # Check if hours of operation are defined
        if (self.workplace not in self.app_data or
            'hours_of_operation' not in self.app_data[self.workplace] or
            not any(blocks for blocks in self.app_data[self.workplace]['hours_of_operation'].values())):
            QMessageBox.warning(self, "Warning", "No hours of operation defined for this workplace. Please define them first.")
            return
        
        # Show dialog to configure schedule generation
        dialog = QDialog(self)
        dialog.setWindowTitle("Generate Schedule")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # Max hours per worker
        max_hours_spin = QSpinBox()
        max_hours_spin.setRange(1, 40)
        max_hours_spin.setValue(20)
        form_layout.addRow("Max Hours Per Worker:", max_hours_spin)
        
        # Max workers per shift
        max_workers_spin = QSpinBox()
        max_workers_spin.setRange(1, 10)
        max_workers_spin.setValue(1)
        form_layout.addRow("Max Workers Per Shift:", max_workers_spin)
        
        layout.addLayout(form_layout)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        generate_btn = StyleHelper.create_button("Generate")
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        
        buttons_layout.addWidget(generate_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
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
            # Load worker data
            file_path = os.path.join("workplaces", f"{self.workplace}.xlsx")
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            workers = []
            for _, row in df.iterrows():
                availability = {day: str(row.get(day, "No")).strip().lower() in ['yes', 'y', 'true'] for day in DAYS}
                not_avail_column = next((col for col in row.index if 'not avail' in col.lower()), None)
                not_available = parse_not_available(row.get(not_avail_column, "")) if not_avail_column else {}
                
                workers.append({
                    "first_name": row.get("First Name", "").strip(),
                    "last_name": row.get("Last Name", "").strip(),
                    "email": row.get("Email", "").strip(),
                    "work_study": str(row.get("Work Study", "")).strip().lower() in ['yes', 'y', 'true'],
                    "availability": availability,
                    "not_available": not_available
                })
            
            # Get hours of operation
            hours_of_operation = self.app_data[self.workplace]['hours_of_operation']
            
            # Generate schedule
            schedule, assigned_hours, low_hour_workers = create_shifts_from_availability(
                hours_of_operation,
                workers,
                self.workplace,
                max_hours_per_worker,
                max_workers_per_shift
            )
            
            # Close dialog
            dialog.accept()
            
            # Show schedule
            self.show_schedule_dialog(schedule, assigned_hours, low_hour_workers)
            
        except Exception as e:
            logging.error(f"Error generating schedule: {str(e)}")
            QMessageBox.critical(dialog, "Error", f"Error generating schedule: {str(e)}")
    
    def show_schedule_dialog(self, schedule, assigned_hours, low_hour_workers):
        """Show dialog with generated schedule"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Generated Schedule")
        dialog.setMinimumWidth(800)
        dialog.setMinimumHeight(600)
        
        layout = QVBoxLayout()
        
        # Schedule tabs
        tabs = QTabWidget()
        
        # Schedule tab
        schedule_tab = QWidget()
        schedule_layout = QVBoxLayout()
        
        # Create schedule table
        for day, shifts in schedule.items():
            if not shifts:
                continue
            
            day_group = QGroupBox(day)
            day_layout = QVBoxLayout()
            
            table = QTableWidget()
            table.setColumnCount(3)
            table.setHorizontalHeaderLabels(["Start", "End", "Assigned"])
            table.setRowCount(len(shifts))
            
            for i, shift in enumerate(shifts):
                # Start
                start_item = QTableWidgetItem(format_time_ampm(shift['start']))
                table.setItem(i, 0, start_item)
                
                # End
                end_item = QTableWidgetItem(format_time_ampm(shift['end']))
                table.setItem(i, 1, end_item)
                
                # Assigned
                assigned_item = QTableWidgetItem(", ".join(shift['assigned']))
                if "Unfilled" in shift['assigned']:
                    assigned_item.setBackground(QColor(255, 200, 200))
                table.setItem(i, 2, assigned_item)
            
            table.resizeColumnsToContents()
            day_layout.addWidget(table)
            day_group.setLayout(day_layout)
            
            schedule_layout.addWidget(day_group)
        
        schedule_tab.setLayout(schedule_layout)
        
        # Hours tab
        hours_tab = QWidget()
        hours_layout = QVBoxLayout()
        
        hours_table = QTableWidget()
        hours_table.setColumnCount(2)
        hours_table.setHorizontalHeaderLabels(["Worker", "Hours"])
        
        # Sort workers by hours (descending)
        sorted_workers = sorted(assigned_hours.items(), key=lambda x: x[1], reverse=True)
        
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
            if hours < 4:
                hours_item.setBackground(QColor(255, 200, 200))
            hours_table.setItem(i, 1, hours_item)
        
        hours_table.resizeColumnsToContents()
        hours_layout.addWidget(hours_table)
        
        # Low hour workers warning
        if low_hour_workers:
            warning_label = QLabel(f"Warning: The following workers have fewer than 4 hours: {', '.join(low_hour_workers)}")
            warning_label.setStyleSheet("color: red;")
            hours_layout.addWidget(warning_label)
        
        hours_tab.setLayout(hours_layout)
        
        # Add tabs
        tabs.addTab(schedule_tab, "Schedule")
        tabs.addTab(hours_tab, "Worker Hours")
        
        layout.addWidget(tabs)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        save_btn = StyleHelper.create_button("Save Schedule")
        email_btn = StyleHelper.create_button("Email Schedule")
        close_btn = StyleHelper.create_button("Close", primary=False)
        
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(email_btn)
        buttons_layout.addWidget(close_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
        save_btn.clicked.connect(lambda: self.save_schedule(dialog, schedule))
        email_btn.clicked.connect(lambda: self.email_schedule_dialog(schedule))
        close_btn.clicked.connect(dialog.reject)
        
        dialog.exec_()
    
    def get_workers(self):
        """Get workers from Excel file"""
        file_path = os.path.join("workplaces", f"{self.workplace}.xlsx")
        
        if not os.path.exists(file_path):
            return []
        
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            workers = []
            for _, row in df.iterrows():
                workers.append({
                    "first_name": row.get("First Name", "").strip(),
                    "last_name": row.get("Last Name", "").strip(),
                    "email": row.get("Email", "").strip(),
                    "work_study": str(row.get("Work Study", "")).strip().lower() in ['yes', 'y', 'true']
                })
            
            return workers
        
        except Exception as e:
            logging.error(f"Error getting workers: {str(e)}")
            return []
    
    def save_schedule(self, dialog, schedule):
        """Save schedule to file"""
        try:
            # Create save path
            save_path = os.path.join("saved_schedules", f"{self.workplace}_current.json")
            
            # Save schedule
            with open(save_path, "w") as f:
                json.dump(schedule, f, indent=4)
            
            QMessageBox.information(dialog, "Success", "Schedule saved successfully.")
            
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
        
        # Sender email
        sender_email_input = QLineEdit()
        form_layout.addRow("Sender Email:", sender_email_input)
        
        # Sender password
        sender_password_input = QLineEdit()
        sender_password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Sender Password:", sender_password_input)
        
        # Recipients
        recipients_input = QTextEdit()
        recipients_input.setPlaceholderText("Enter email addresses, one per line")
        
        # Add worker emails
        workers = self.get_workers()
        for worker in workers:
            if worker['email']:
                recipients_input.append(worker['email'])
        
        form_layout.addRow("Recipients:", recipients_input)
        
        layout.addLayout(form_layout)
        
        # Buttons
        buttons_layout = QHBoxLayout()
        
        send_btn = StyleHelper.create_button("Send")
        cancel_btn = StyleHelper.create_button("Cancel", primary=False)
        
        buttons_layout.addWidget(send_btn)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        
        # Connect buttons
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
            # Send email
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
    
    def view_current_schedule(self):
        """View current saved schedule"""
        # Check if schedule exists
        save_path = os.path.join("saved_schedules", f"{self.workplace}_current.json")
        
        if not os.path.exists(save_path):
            QMessageBox.warning(self, "Warning", "No saved schedule found for this workplace.")
            return
        
        try:
            # Load schedule
            with open(save_path, "r") as f:
                schedule = json.load(f)
            
            # Show schedule
            self.show_schedule_dialog(schedule, {}, [])
            
        except Exception as e:
            logging.error(f"Error viewing schedule: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error viewing schedule: {str(e)}")

class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(1000, 700)
        
        # Set style
        self.setStyleSheet(StyleHelper.get_main_style())
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Header
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
        
        # Tabs
        tabs = QTabWidget()
        
        # Add workplace tabs
        tabs.addTab(WorkplaceTab("esports_lounge"), "eSports Lounge")
        tabs.addTab(WorkplaceTab("esports_arena"), "eSports Arena")
        tabs.addTab(WorkplaceTab("it_service_center"), "IT Service Center")
        
        main_layout.addWidget(tabs)
        
        # Show window
        self.show()

# Main function
def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
