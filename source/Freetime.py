import requests
#from ics import Calendar
from datetime import datetime, timedelta
import pytz
import pyperclip
from typing import List, Dict, Set
from datetime import date
import pystray
#from PIL import Image
import threading
import time
import pyautogui
import tempfile
import os
import json
import logging
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import sys
from pathlib import Path
import platform
import subprocess
import tkinter.font as tk_font
#from pynput import keyboard
import keyboard
from PIL import Image, ImageTk
from icalendar import Calendar, vDatetime
from dateutil import rrule

version_string = "0.8"
# Changelog

# v0.8 7/3/25
# Added option to ignore allday and multi-day events

# v0.7 3/3/25
# Switched to text replacement rather than hotkey

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def get_app_directory():
    """Get the directory where the app is running from"""
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle (compiled)
        return Path(sys.executable).parent
    else:
        # If the application is run from a Python interpreter
        return Path(__file__).parent

class NewestFirstLogHandler(logging.FileHandler):
    MAX_ENTRIES = 1000

    def emit(self, record):
        try:
            # Read existing logs
            lines = []
            if os.path.exists(self.baseFilename):
                with open(self.baseFilename, 'r') as f:
                    lines = f.readlines()

            # Add new log entry at the beginning
            lines.insert(0, self.format(record) + '\n')

            # Keep only the most recent MAX_ENTRIES
            lines = lines[:self.MAX_ENTRIES]

            # Write back to file
            with open(self.baseFilename, 'w') as f:
                f.writelines(lines)
        except Exception:
            self.handleError(record)

# Set up app paths
APP_DIR = get_app_directory()
SETTINGS_FILE = APP_DIR / 'calendar_settings.json'
LOG_FILE = APP_DIR / 'freetime.log'

try:
    # Test write permissions by touching the files
    SETTINGS_FILE.touch(exist_ok=True)
    LOG_FILE.touch(exist_ok=True)
except PermissionError:
    # If we can't write to the app directory, fall back to user's temp directory
    temp_dir = Path(tempfile.gettempdir()) / 'CalendarApp'
    temp_dir.mkdir(exist_ok=True)
    SETTINGS_FILE = temp_dir / 'calendar_settings.json'
    LOG_FILE = temp_dir / 'calendar_app.log'
    logging.warning(f"Cannot write to app directory. Using temporary directory: {temp_dir}")

# Set up logging
logging.basicConfig(
    level=logging.INFO, # should be .INFO
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        NewestFirstLogHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)

class AboutWindow:
    def __init__(self, root, icon_image):
        self.window = tk.Toplevel()
        self.window.title("Freetime")
        self.window.geometry("300x230")

        # Make window non-resizable
        self.window.resizable(False, False)

        # Center the window
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')

        # Set window icon
        try:
            icon_photo = ImageTk.PhotoImage(icon_image)
            self.window.iconphoto(True, icon_photo)
            self.icon_photo = icon_photo  # Keep a reference
        except Exception as e:
            logging.error(f"Error setting window icon: {e}")

        # Create main frame with padding
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(expand=True, fill='both')

        # Create inner frame for content
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(expand=True)

        # Display logo
        try:
            # Resize image to fit window (adjust size as needed)
            logo_size = (100, 100)
            logo_image = icon_image.resize(logo_size, Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(content_frame, image=photo)
            logo_label.image = photo  # Keep a reference
            logo_label.pack(pady=(0, 10))
        except Exception as e:
            logging.error(f"Error displaying logo: {e}")

        # Version text
        version_label = ttk.Label(content_frame, text=f"Freetime {version_string} by Vince Polito")
        version_label.pack(pady=(0, 10))

        # GitHub link
        link_label = ttk.Label(content_frame, text="GitHub Page",
                               cursor="hand2", foreground="blue")
        link_font = tk_font.Font(
            family=link_label.cget("font"),
            size=10,
            underline=True
        )
        link_label.configure(font=link_font)
        link_label.pack()

        def open_github(event):
            try:
                import webbrowser
                webbrowser.open("https://github.com/vince-p/freetime", new=1)
            except Exception as e:
                logging.error(f"Error opening GitHub link: {e}")
                messagebox.showerror("Error",
                                     "Could not open link. Please visit https://github.com/vince-p/freetime manually.")

        link_label.bind("<Button-1>", open_github)

        # Keep window on top
        self.window.lift()
        self.window.focus_force()


class SettingsWindow:
    def __init__(self, app, root):
        self.app = app
        self.window = tk.Toplevel(root)
        self.window.title("FreeTime Settings")
        self.window.geometry("500x740")

        # Set the window icon
        try:
            # Convert PIL image to PhotoImage for Tkinter
            icon_photo = ImageTk.PhotoImage(self.app.icon_image)
            self.window.iconphoto(True, icon_photo)
            # Keep a reference to prevent garbage collection
            self.icon_photo = icon_photo
        except Exception as e:
            logging.error(f"Error setting window icon: {e}")

        # Configure column weight to allow frames to expand
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)

        self.setup_ui()

        self.window.update_idletasks()
        self.window.deiconify()
        self.window.focus_force()

    def open_log_file(self, event):
        """Open the log file with the default system application"""
        try:
            if platform.system() == 'Windows':
                os.startfile(LOG_FILE)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', LOG_FILE])
            else:  # Linux
                subprocess.run(['xdg-open', LOG_FILE])
        except Exception as e:
            logging.error(f"Failed to open log file: {e}")
            messagebox.showerror("Error", f"Could not open log file: {e}")

    def add_calendar_url(self):
        """Prompt user for new calendar URL"""
        url = simpledialog.askstring("Add Calendar", "Enter calendar URL:", parent=self.window)
        if url:
            if url.startswith(('http://', 'https://')):
                self.url_listbox.insert(tk.END, url)
            else:
                messagebox.showerror("Invalid URL", "URL must start with http:// or https://")

    def remove_calendar_url(self):
        """Remove selected calendar URL"""
        selection = self.url_listbox.curselection()
        if selection:
            self.url_listbox.delete(selection)

    def setup_ui(self):
        # Create main frame with padding
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)

        # Define clear_status function at the beginning so it's in scope
        def clear_status(*args):
            # Track changes in the new checkbox
            self.ignore_all_day_events_var.trace_add("write", clear_status)

            """Clear the status message when any setting is changed"""
            if self.status_label.cget("text") == "Settings saved":
                self.status_label.config(text="")

        # Create bold font for group labels
        current_font = tk_font.Font(font='TkDefaultFont')
        bold_font = tk_font.Font(
            family=current_font.cget("family"),
            size=current_font.cget("size"),
            weight="bold"
        )

        # Create style for bold label frames
        style = ttk.Style()
        style.configure('Bold.TLabelframe.Label', font=bold_font)

        current_row = 0

        # Calendar Settings Section (with bold label)
        calendar_frame = ttk.LabelFrame(main_frame, text="Calendar Settings", padding="5", style='Bold.TLabelframe')
        calendar_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5, padx=5)
        calendar_frame.columnconfigure(0, weight=1)

        # URL Listbox
        ttk.Label(calendar_frame, text="Calendar URLs:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.url_listbox = tk.Listbox(calendar_frame, height=5)
        self.url_listbox.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))

        # Add scrollbar to URL Listbox
        url_scrollbar = ttk.Scrollbar(calendar_frame, orient=tk.VERTICAL, command=self.url_listbox.yview)
        url_scrollbar.grid(row=1, column=3, sticky=(tk.N, tk.S))
        self.url_listbox.configure(yscrollcommand=url_scrollbar.set)

        # Populate URL Listbox
        for url in self.app.calendar_urls:
            self.url_listbox.insert(tk.END, url)

        # URL Buttons
        ttk.Button(calendar_frame, text="Add", command=self.add_calendar_url).grid(
            row=2, column=0, sticky=tk.W, pady=5)
        ttk.Button(calendar_frame, text="Remove", command=self.remove_calendar_url).grid(
            row=2, column=1, sticky=tk.W, pady=5)

        # Custom Text
        ttk.Label(calendar_frame, text="Custom Text:").grid(row=3, column=0, sticky=tk.W, pady=(10, 0))
        self.custom_text_var = tk.StringVar(value=self.app.custom_text)
        self.custom_text_entry = ttk.Entry(calendar_frame, textvariable=self.custom_text_var)
        self.custom_text_entry.grid(row=4, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 5))

        current_row += 1

        # Time Settings Section (with bold label)
        time_frame = ttk.LabelFrame(main_frame, text="Time Settings", padding="5", style='Bold.TLabelframe')
        time_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)
        time_frame.columnconfigure(1, weight=1)

        # Working Hours
        ttk.Label(time_frame, text="Meeting Hours:").grid(row=0, column=0, sticky=tk.W)
        hours_frame = ttk.Frame(time_frame)
        hours_frame.grid(row=0, column=1, sticky=tk.W)

        self.start_hour_var = tk.IntVar(value=self.app.start_of_day)
        self.end_hour_var = tk.IntVar(value=self.app.end_of_day)

        ttk.Label(hours_frame, text="Start:").grid(row=0, column=0, sticky=tk.W)
        ttk.Spinbox(hours_frame, from_=0, to=23, width=5, textvariable=self.start_hour_var).grid(
            row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(hours_frame, text="End:").grid(row=0, column=2, sticky=tk.W, padx=10)
        ttk.Spinbox(hours_frame, from_=0, to=23, width=5, textvariable=self.end_hour_var).grid(
            row=0, column=3, sticky=tk.W, padx=5)

        # Lookahead Days
        ttk.Label(time_frame, text="Lookahead Days:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.lookahead_var = tk.IntVar(value=self.app.lookahead_days)
        ttk.Spinbox(time_frame, from_=1, to=30, width=5, textvariable=self.lookahead_var).grid(
            row=1, column=1, sticky=tk.W)

        # Include Current Day
        self.include_current_day_var = tk.BooleanVar(value=self.app.include_current_day)
        ttk.Checkbutton(time_frame, text="Include Current Day", variable=self.include_current_day_var).grid(
            row=2, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Exclude Weekends
        self.exclude_weekends_var = tk.BooleanVar(value=self.app.exclude_weekends)
        ttk.Checkbutton(time_frame, text="Exclude Weekends", variable=self.exclude_weekends_var).grid(
            row=3, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Ignore all-day events
        self.ignore_all_day_events_var = tk.BooleanVar(value=self.app.ignore_all_day_events)
        ttk.Checkbutton(time_frame, text="Ignore all-day and multi-day events",
                        variable=self.ignore_all_day_events_var).grid(
            row=4, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Timezone
        ttk.Label(time_frame, text="Timezone:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.timezone_var = tk.StringVar(value=str(self.app.local_tz))
        self.timezone_combo = ttk.Combobox(time_frame, textvariable=self.timezone_var)
        self.timezone_combo['values'] = pytz.all_timezones
        self.timezone_combo.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=5)

        current_row += 1

        # App Settings Section (with bold label)
        app_frame = ttk.LabelFrame(main_frame, text="App Settings", padding="5", style='Bold.TLabelframe')
        app_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)
        app_frame.columnconfigure(1, weight=1)

        # Update Interval
        ttk.Label(app_frame, text="Update Interval (minutes):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.interval_var = tk.IntVar(value=self.app.update_interval // 60)
        ttk.Spinbox(app_frame, from_=1, to=60, width=5, textvariable=self.interval_var).grid(
            row=0, column=1, sticky=tk.W, pady=5)

        # Trigger Text (replacing Hotkey)
        ttk.Label(app_frame, text="Trigger Text:").grid(row=1, column=0, sticky=tk.W)
        self.trigger_text_var = tk.StringVar(value=getattr(self.app, 'trigger_pattern', ":ttt"))
        self.trigger_text_entry = ttk.Entry(app_frame, textvariable=self.trigger_text_var)
        self.trigger_text_entry.grid(row=1, column=1, sticky=(tk.W, tk.E))

        # Add a help label explaining the trigger
        #ttk.Label(app_frame, text="Type this text anywhere to insert your free time",
        #          font=('', 9), foreground='gray').grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5)

        # Startup checkbox
        self.startup_var = tk.BooleanVar(value=self.check_startup())
        startup_cb = ttk.Checkbutton(app_frame, text="Run at startup",
                                     variable=self.startup_var,
                                     command=self.toggle_startup)
        startup_cb.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)

        # Bottom Section - Create a frame for the bottom elements
        # Add an empty row with weight to push the bottom frame down
        main_frame.rowconfigure(current_row, weight=1)  # Give weight to push content down
        current_row += 1

        # Now create the bottom frame in the next row
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=current_row, column=0, pady=10, sticky=(tk.W, tk.E, tk.S))  # Add sticky=tk.S
        bottom_frame.columnconfigure(1, weight=1)  # Make middle column expand

        # Log File Link (left aligned)
        log_label = ttk.Label(bottom_frame, text="View log", cursor="hand2")
        log_label.grid(row=0, column=0, sticky=tk.W)

        # Get the font used by other ttk elements
        default_font = ttk.Style().lookup('TLabel', 'font')
        if isinstance(default_font, str):
            default_font = tk_font.nametofont(default_font)

        link_font = tk_font.Font(
            family=default_font.cget("family"),
            size=default_font.cget("size") - 1,
            weight=default_font.cget("weight")
        )
        link_font.configure(underline=True)
        log_label.configure(font=link_font)
        log_label.bind("<Button-1>", self.open_log_file)

        # Status Label (center aligned)
        self.status_label = ttk.Label(bottom_frame, text="", foreground="green")
        self.status_label.grid(row=0, column=1)

        # Show startup message if no URLs exist
        if not self.app.calendar_urls:
            # Create a frame for the message and link
            msg_frame = ttk.Frame(bottom_frame)
            msg_frame.grid(row=0, column=1)

            # Add the text message
            msg_label = ttk.Label(msg_frame, text="Enter an ical url to start. ", foreground="blue")
            msg_label.grid(row=0, column=0)

            # Create and style the hyperlink
            link_label = ttk.Label(msg_frame, text="See here", cursor="hand2", foreground="blue")
            link_font = tk_font.Font(
                family=default_font.cget("family"),
                size=default_font.cget("size"),
                underline=True
            )
            link_label.configure(font=link_font)
            link_label.grid(row=0, column=1)

            # Add click handler for the link
            def open_link(event):
                import webbrowser
                webbrowser.open("https://github.com/vince-p/freetime")

            link_label.bind("<Button-1>", open_link)

        # Save Button (right aligned)
        ttk.Button(bottom_frame, text="Save", command=self.save_settings).grid(
            row=0, column=2, sticky=tk.E)

        # Add change trackers for the trigger text
        self.trigger_text_var.trace_add("write", clear_status)

        # Add change trackers to all input elements
        def clear_status(*args):
            """Clear the status message when any setting is changed"""
            if self.status_label.cget("text") == "Settings saved":
                self.status_label.config(text="")

            # Add change trackers to all input elements
            # Track changes in entry widgets and listbox
            self.url_listbox.bind('<<ListboxSelect>>', clear_status)
            self.custom_text_var.trace_add("write", clear_status)

            # Track changes in spinboxes
            self.start_hour_var.trace_add("write", clear_status)
            self.end_hour_var.trace_add("write", clear_status)
            self.lookahead_var.trace_add("write", clear_status)
            self.interval_var.trace_add("write", clear_status)

            # Track changes in checkboxes
            self.exclude_weekends_var.trace_add("write", clear_status)
            self.include_current_day_var.trace_add("write", clear_status)

            # Track changes in timezone combobox
            self.timezone_combo.bind('<<ComboboxSelected>>', clear_status)
            self.timezone_var.trace_add("write", clear_status)

            # Track changes in trigger text
            self.trigger_text_var.trace_add("write", clear_status)

        # Also track URL additions and removals
        def on_url_change():
            clear_status()
            self.url_listbox.bind('<<ListboxSelect>>', clear_status)  # Rebind after changes

        # Update the add_calendar_url and remove_calendar_url methods
        def add_calendar_url():
            url = simpledialog.askstring("Add Calendar", "Enter calendar URL:", parent=self.window)
            if url:
                if url.startswith(('http://', 'https://')):
                    self.url_listbox.insert(tk.END, url)
                    on_url_change()
                else:
                    messagebox.showerror("Invalid URL", "URL must start with http:// or https://")

        def remove_calendar_url():
            selection = self.url_listbox.curselection()
            if selection:
                self.url_listbox.delete(selection)
                on_url_change()

        # Update the button commands
        self.add_button = ttk.Button(calendar_frame, text="Add", command=add_calendar_url)
        self.add_button.grid(row=2, column=0, sticky=tk.W, pady=5)

        self.remove_button = ttk.Button(calendar_frame, text="Remove", command=remove_calendar_url)
        self.remove_button.grid(row=2, column=1, sticky=tk.W, pady=5)

    def save_settings(self):
        try:
            # Get new URLs from listbox
            new_urls = list(self.url_listbox.get(0, tk.END))
            urls_changed = set(new_urls) != set(self.app.calendar_urls)

            # Check if URLs exist
            if not new_urls:
                self.status_label.config(
                    text="Enter at least one ical url to get started",
                    foreground="blue"
                )
                return

            # Store old trigger pattern
            old_trigger = getattr(self.app, 'trigger_pattern', ":ttt")
            new_trigger = self.trigger_text_var.get()

            # Update all settings
            self.app.calendar_urls = new_urls
            self.app.local_tz = pytz.timezone(self.timezone_var.get())
            self.app.start_of_day = self.start_hour_var.get()
            self.app.end_of_day = self.end_hour_var.get()
            self.app.lookahead_days = self.lookahead_var.get()
            self.app.update_interval = self.interval_var.get() * 60
            self.app.custom_text = self.custom_text_var.get()
            self.app.exclude_weekends = self.exclude_weekends_var.get()
            self.app.include_current_day = self.include_current_day_var.get()
            self.app.ignore_all_day_events = self.ignore_all_day_events_var.get()

            # Handle trigger pattern change
            if new_trigger != old_trigger:
                self.app.trigger_pattern = new_trigger
                self.app.setup_hotkey()  # This will reset the text pattern detection
                logging.info("Trigger text changed")

            # Save settings to file
            self.app.save_settings()

            # If URLs changed, clear cache
            if urls_changed:
                logging.info("Calendar URLs changed, clearing cache...")
                self.app.clear_cache()

            logging.info("Settings updated and saved")
            self.status_label.config(text="Settings saved", foreground="green")
            self.app.update_free_slots()

        except Exception as e:
            error_msg = f"Failed to save settings: {str(e)}"
            logging.error(error_msg)
            self.status_label.config(text=error_msg, foreground="red")



    def check_startup(self):
        """Check if app is set to run at startup"""
        if not getattr(sys, 'frozen', False):
            return False

        try:
            if platform.system() == 'Windows':
                import winreg
                key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                                        winreg.KEY_READ) as key:
                        winreg.QueryValueEx(key, "Freetime")
                        return True
                except WindowsError:
                    return False

            elif platform.system() == 'Darwin':  # macOS
                plist_path = os.path.expanduser(
                    '~/Library/LaunchAgents/com.freetime.startup.plist')
                return os.path.exists(plist_path)

            else:  # Linux
                autostart_path = os.path.expanduser(
                    '~/.config/autostart/freetime.desktop')
                return os.path.exists(autostart_path)

        except Exception as e:
            logging.error(f"Error checking startup status: {e}")
            return False

    def toggle_startup(self):
        """Toggle startup status"""
        if not getattr(sys, 'frozen', False):
            messagebox.showinfo("Not Available",
                                "Start-up option is only available when running as a compiled program.")
            self.startup_var.set(False)  # Reset checkbox
            return

        try:
            if platform.system() == 'Windows':
                self._toggle_startup_windows()
            elif platform.system() == 'Darwin':  # macOS
                self._toggle_startup_macos()
            else:  # Linux
                self._toggle_startup_linux()
        except Exception as e:
            logging.error(f"Error toggling startup: {e}")
            self.startup_var.set(not self.startup_var.get())  # Revert checkbox
            messagebox.showerror("Error",
                                 f"Failed to modify startup settings: {str(e)}")

    def _toggle_startup_windows(self):
        """Toggle startup on Windows"""
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_path = sys.executable

        try:
            if self.startup_var.get():  # Add to startup
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                                    winreg.KEY_WRITE) as key:
                    winreg.SetValueEx(key, "Freetime", 0, winreg.REG_SZ, app_path)
            else:  # Remove from startup
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0,
                                    winreg.KEY_WRITE) as key:
                    try:
                        winreg.DeleteValue(key, "Freetime")
                    except WindowsError:
                        pass
        except Exception as e:
            raise Exception(f"Failed to modify Windows registry: {e}")

    def _toggle_startup_macos(self):
        """Toggle startup on macOS"""
        plist_path = os.path.expanduser(
            '~/Library/LaunchAgents/com.freetime.startup.plist')
        app_path = sys.executable

        plist_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.freetime.startup</string>
    <key>ProgramArguments</key>
    <array>
        <string>{app_path}</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
</dict>
</plist>"""

        try:
            if self.startup_var.get():  # Add to startup
                os.makedirs(os.path.dirname(plist_path), exist_ok=True)
                with open(plist_path, 'w') as f:
                    f.write(plist_content)
                os.chmod(plist_path, 0o644)
                subprocess.run(['launchctl', 'load', plist_path])
            else:  # Remove from startup
                if os.path.exists(plist_path):
                    subprocess.run(['launchctl', 'unload', plist_path])
                    os.remove(plist_path)
        except Exception as e:
            raise Exception(f"Failed to modify macOS startup: {e}")

    def _toggle_startup_linux(self):
        """Toggle startup on Linux"""
        autostart_dir = os.path.expanduser('~/.config/autostart')
        desktop_file = os.path.join(autostart_dir, 'freetime.desktop')
        app_path = sys.executable

        desktop_content = f"""[Desktop Entry]
Type=Application
Exec={app_path}
Hidden=false
NoDisplay=false
X-GNOME-Autostart-enabled=true
Name=Freetime
Comment=Calendar Free Time Finder"""

        try:
            if self.startup_var.get():  # Add to startup
                os.makedirs(autostart_dir, exist_ok=True)
                with open(desktop_file, 'w') as f:
                    f.write(desktop_content)
                os.chmod(desktop_file, 0o755)
            else:  # Remove from startup
                if os.path.exists(desktop_file):
                    os.remove(desktop_file)
        except Exception as e:
            raise Exception(f"Failed to modify Linux startup: {e}")



class CalendarApp:
    def __init__(self):
        self.settings_file = SETTINGS_FILE
        self.cache_file = APP_DIR / 'calendar_cache.json'

        # Load icon first
        self.icon_image = self.load_icon()

        # Log startup
        logging.info("FreeTime application starting...")
        logging.info(f"App directory: {APP_DIR}")
        logging.info(f"Settings file: {self.settings_file}")
        logging.info(f"Cache file: {self.cache_file}")

        self.load_settings()
        self.cached_free_slots = None
        self.load_cache()
        self.settings_window = None

        # Initialize the root window
        self.root = tk.Tk()
        self.root.withdraw()
        self.icon = None

        self.original_clipboard = None

        self.about_window = None

    from pynput import keyboard

    def setup_hotkey(self):
        """Setup the text pattern detector"""
        try:
            logging.info("Starting text pattern detection setup...")

            # The trigger pattern to look for
            self.trigger_pattern = getattr(self, 'trigger_pattern', ":ttt")  # Default if not set
            self.char_buffer = ""
            self.buffer_max_size = 20  # Keep this larger than the trigger pattern
            self.is_pasting = False

            def on_key(event):
                try:
                    if self.is_pasting:
                        return

                    # Skip if we're typing in the settings window
                    if self.settings_window is not None and tk.Toplevel.winfo_exists(self.settings_window.window):
                        active_widget = self.settings_window.window.focus_get()
                        if active_widget and isinstance(active_widget, (ttk.Entry, tk.Entry, ttk.Combobox)):
                            return

                    # Add character to buffer
                    if event.event_type == keyboard.KEY_DOWN and hasattr(event, 'name'):
                        if event.name == 'space':
                            self.char_buffer += ' '
                        elif event.name == 'enter':
                            self.char_buffer = ""  # Clear buffer on enter
                        elif event.name == 'backspace':
                            if self.char_buffer:
                                self.char_buffer = self.char_buffer[:-1]
                        elif len(event.name) == 1:  # Only add printable characters
                            self.char_buffer += event.name

                        # Keep buffer at manageable size
                        if len(self.char_buffer) > self.buffer_max_size:
                            self.char_buffer = self.char_buffer[-self.buffer_max_size:]

                        # Check if trigger pattern is in buffer
                        if self.trigger_pattern in self.char_buffer:
                            logging.info(f"Trigger pattern detected: {self.trigger_pattern}")

                            # Delete the trigger text by simulating backspace presses
                            for _ in range(len(self.trigger_pattern)):
                                keyboard.send('backspace')

                            # Clear buffer and trigger paste
                            self.char_buffer = ""
                            self.trigger_paste()

                except Exception as e:
                    logging.error(f"Error in key handler: {e}", exc_info=True)

            # Remove any existing keyboard hooks
            keyboard.unhook_all()

            # Setup the new keyboard listener
            keyboard.hook(on_key)

            logging.info(f"Text pattern detection set up for: {self.trigger_pattern}")

        except Exception as e:
            logging.error(f"Failed to setup text pattern detection: {e}", exc_info=True)

    def trigger_paste(self):
        """Separate method to handle paste operation"""
        try:
            self.is_pasting = True
            self.paste_free_slots()
        finally:
            self.is_pasting = False

    def debug_keyboard_state(self):
        """Debug helper to print current keyboard state"""
        logging.info("=== Keyboard State Debug ===")
        logging.info(f"Hotkey keys: {self.hotkey_keys}")
        logging.info(f"Current keys: {self.current_keys}")
        #logging.info(f"Time since last trigger: {time.time() - self.last_trigger_time}")
        logging.info("==========================")

    def restore_clipboard(self):
        """Restore original clipboard content"""
        try:
            if self.original_clipboard is not None:
                pyperclip.copy(self.original_clipboard)
                self.original_clipboard = None
        except Exception as e:
            logging.error(f"Error restoring clipboard: {e}")

    def clear_cache(self):
        """Clear all cached data"""
        try:
            if os.path.exists(self.cache_file):
                os.remove(self.cache_file)
            self.cached_free_slots = None
            logging.info("Cache cleared successfully")
        except Exception as e:
            logging.error(f"Error clearing cache: {e}")

    def load_settings(self):
        """Load settings from file or use defaults"""
        defaults = {
            'calendar_urls': [],
            'timezone': "Australia/Sydney",
            'start_of_day': 9,
            'end_of_day': 16,
            'lookahead_days': 7,
            'include_current_day': False,
            'exclude_weekends': True,
            'update_interval': 300,
            'trigger_pattern': ":ttt",
            'custom_text': "I'm free at the following times...",
            'ignore_all_day_events': True  # New setting, default to True
        }

        try:
            logging.info(f"Attempting to load settings from: {self.settings_file}")

            if self.settings_file.exists():
                with open(self.settings_file, 'r') as f:
                    file_content = f.read()
                    logging.info(f"Settings file content: {file_content}")

                    # Check if file is empty
                    if not file_content.strip():
                        logging.warning("Settings file exists but is empty. Using defaults.")
                        settings = {}
                    else:
                        try:
                            settings = json.loads(file_content)
                        except json.JSONDecodeError as e:
                            logging.error(f"Failed to parse settings JSON: {e}")
                            settings = {}

                    # Apply settings from file or use defaults
                    for key, default_value in defaults.items():
                        if key == 'timezone':
                            timezone_str = settings.get(key, default_value)
                            try:
                                self.local_tz = pytz.timezone(timezone_str)
                                logging.info(f"Loaded timezone: {timezone_str}")
                            except Exception as e:
                                logging.error(f"Invalid timezone {timezone_str}: {e}, using default")
                                self.local_tz = pytz.timezone(default_value)
                        else:
                            value = settings.get(key, default_value)
                            setattr(self, key, value)
                            logging.info(f"Loaded setting {key}: {value}")

                    # For backward compatibility - map hotkey to trigger_pattern if needed
                    if 'hotkey' in settings and not hasattr(self, 'trigger_pattern'):
                        self.trigger_pattern = defaults['trigger_pattern']
                        logging.info("Using default trigger pattern due to legacy settings format")

                    logging.info("Settings loaded successfully from file")
            else:
                logging.info("Settings file not found, using defaults")
                for key, value in defaults.items():
                    if key == 'timezone':
                        self.local_tz = pytz.timezone(value)
                    else:
                        setattr(self, key, value)
                    logging.info(f"Setting default {key}: {value}")
        except Exception as e:
            logging.error(f"Error loading settings: {e}", exc_info=True)
            logging.info("Using default settings due to error")
            for key, value in defaults.items():
                if key == 'timezone':
                    self.local_tz = pytz.timezone(value)
                else:
                    setattr(self, key, value)

    def save_settings(self):
        """Save current settings to file"""
        try:
            settings = {
                'calendar_urls': self.calendar_urls,
                'timezone': str(self.local_tz),
                'start_of_day': self.start_of_day,
                'end_of_day': self.end_of_day,
                'lookahead_days': self.lookahead_days,
                'include_current_day': self.include_current_day,
                'exclude_weekends': self.exclude_weekends,
                'update_interval': self.update_interval,
                'trigger_pattern': self.trigger_pattern,
                'ignore_all_day_events': self.ignore_all_day_events,
                'custom_text': self.custom_text  # Make sure to save custom_text too
            }

            # Log what we're about to save
            logging.info(f"Saving settings to: {self.settings_file}")
            logging.info(f"Settings data: {settings}")

            # Create directory if it doesn't exist
            self.settings_file.parent.mkdir(parents=True, exist_ok=True)

            with open(self.settings_file, 'w') as f:
                json_data = json.dumps(settings, indent=2)
                f.write(json_data)

            logging.info("Settings saved successfully")

            # Verify the file was written correctly
            if self.settings_file.exists():
                file_size = self.settings_file.stat().st_size
                logging.info(f"Settings file size: {file_size} bytes")
            else:
                logging.warning("Settings file was not created!")

        except Exception as e:
            logging.error(f"Error saving settings: {e}", exc_info=True)

    def load_cache(self):
        """Load cached free slots from temporary file"""
        try:
            if self.cache_file.exists():
                with open(self.cache_file, 'r') as f:
                    cached_data = json.load(f)
                    # Convert string dates back to datetime objects
                    self.cached_free_slots = {
                        datetime.strptime(date_str, '%Y-%m-%d').date():
                            [datetime.fromisoformat(slot_str) for slot_str in slots]
                        for date_str, slots in cached_data.items()
                    }
            else:
                self.cached_free_slots = {}
        except Exception as e:
            logging.error(f"Error loading cache: {e}")
            self.cached_free_slots = {}

    def save_cache(self):
        """Save free slots to temporary file"""
        try:
            if self.cached_free_slots is None:
                return

            # Convert datetime objects to strings for JSON serialization
            cache_data = {
                date_obj.isoformat(): [slot.isoformat() for slot in slots]
                for date_obj, slots in self.cached_free_slots.items()
            }
            with open(self.cache_file, 'w') as f:
                json.dump(cache_data, f)
        except Exception as e:
            logging.error(f"Error saving cache: {e}")

    def show_settings(self):
        """Show the settings window"""
        try:
            if self.settings_window is None or not tk.Toplevel.winfo_exists(self.settings_window.window):
                self.settings_window = SettingsWindow(self, self.root)
            else:
                self.settings_window.window.lift()
                self.settings_window.window.focus_force()
        except Exception as e:
            logging.error(f"Error showing settings window: {e}")

    def is_weekend(self, date_obj):
        """Check if the given date is a weekend (Saturday=5 or Sunday=6)"""
        return date_obj.weekday() >= 5

    def fetch_and_parse_calendar(self, url: str) -> Dict[date, Set[datetime]]:
        """Fetches an .ics calendar from a URL and parses it into free one-hour time slots."""
        start_time = time.time()
        try:
            response = requests.get(url)
            response.raise_for_status()

            # Add error checking for HTML response
            if 'text/html' in response.headers.get('content-type', '').lower():
                logging.error(f"URL returned HTML instead of iCal data: {url}")
                return {}

            gcal = Calendar.from_ical(response.text)

            now = datetime.now(self.local_tz)
            free_slots = {}

            # Calculate date range
            start_date = now.date()
            end_date = start_date + timedelta(days=self.lookahead_days)

            # Create a dict to store busy times for each day
            busy_times_by_day = {}

            # Debug: Log calendar processing start
            #logging.info(f"\n=== Processing calendar: {url} ===")

            for component in gcal.walk():
                if component.name == "VEVENT":
                    event_start = component.get('dtstart').dt
                    event_end = component.get('dtend').dt
                    event_summary = str(component.get('summary', 'No Title'))
                    rrule_str = component.get('rrule')

                    # Check if it's an all-day event
                    is_all_day = (isinstance(event_start, date) and not isinstance(event_start, datetime))

                    # Skip this event if it's all-day and the ignore option is enabled
                    if is_all_day and self.ignore_all_day_events:
                        continue

                    # Convert to datetime if date
                    if isinstance(event_start, date) and not isinstance(event_start, datetime):
                        event_start = datetime.combine(event_start, datetime.min.time())
                    if isinstance(event_end, date) and not isinstance(event_end, datetime):
                        event_end = datetime.combine(event_end, datetime.min.time())

                    # Ensure timezone awareness
                    if event_start.tzinfo is None:
                        event_start = self.local_tz.localize(event_start)
                    if event_end.tzinfo is None:
                        event_end = self.local_tz.localize(event_end)

                    # Convert to local timezone if needed
                    event_start = event_start.astimezone(self.local_tz)
                    event_end = event_end.astimezone(self.local_tz)

                    # For multi-day events, create entries for each day in the range
                    if (event_end.date() - event_start.date()).days > 0:
                        # Check if it's a multi-day event to be ignored
                        if self.ignore_all_day_events:
                            continue

                        # Process each day of the multi-day event
                        current_day = event_start.date()
                        end_day = event_end.date()

                        while current_day <= end_day:
                            day_start = datetime.combine(current_day, datetime.min.time())
                            day_start = self.local_tz.localize(day_start)

                            day_end = datetime.combine(current_day, datetime.max.time())
                            day_end = self.local_tz.localize(day_end)

                            # Adjust first and last day times
                            if current_day == event_start.date():
                                day_start = event_start
                            if current_day == end_day:
                                day_end = event_end

                            if current_day not in busy_times_by_day:
                                busy_times_by_day[current_day] = []
                            busy_times_by_day[current_day].append((day_start, day_end, event_summary))

                            current_day += timedelta(days=1)
                    else:
                        # Process regular single-day event as before
                        event_date = event_start.date()
                        if event_date not in busy_times_by_day:
                            busy_times_by_day[event_date] = []
                        busy_times_by_day[event_date].append((event_start, event_end, event_summary))

                    if rrule_str:  # Recurring event
                        # Convert rrule to dateutil rrule
                        rule = rrule.rrulestr(
                            rrule_str.to_ical().decode('utf-8'),
                            dtstart=event_start
                        )

                        # Get all occurrences in our date range
                        event_duration = event_end - event_start
                        for occurrence_start in rule.between(
                                now - event_duration,
                                now + timedelta(days=self.lookahead_days)
                        ):
                            occurrence_end = occurrence_start + event_duration
                            event_date = occurrence_start.date()

                            if event_date not in busy_times_by_day:
                                busy_times_by_day[event_date] = []
                            busy_times_by_day[event_date].append((occurrence_start, occurrence_end, event_summary))
                    else:  # Single event
                        event_date = event_start.date()
                        if event_date not in busy_times_by_day:
                            busy_times_by_day[event_date] = []
                        busy_times_by_day[event_date].append((event_start, event_end, event_summary))

            # Process each day in the range
            for day_offset in range(self.lookahead_days):
                day_start = self.local_tz.localize(
                    datetime.combine(start_date + timedelta(days=day_offset),
                                     datetime.min.time().replace(hour=self.start_of_day)))

                current_date = day_start.date()

                # Skip weekends if exclude_weekends is True
                if self.exclude_weekends and self.is_weekend(current_date):
                    continue

                day_end = self.local_tz.localize(
                    datetime.combine(current_date,
                                     datetime.min.time().replace(hour=self.end_of_day)))

                # Debug: Log day processing
                #logging.info(f"\n=== Processing {current_date.strftime('%A %d %B %Y')} ===")

                # Get busy times for this day and sort them
                busy_times = busy_times_by_day.get(current_date, [])
                busy_times.sort()

                # Debug: Log all events for this day
                #if busy_times:
                #    logging.info(f"Events for {current_date.strftime('%A %d %B')}:")
                #    for busy_start, busy_end, summary in busy_times:
                #        logging.info(
                #            f"  {summary}: {busy_start.strftime('%I:%M %p')} - {busy_end.strftime('%I:%M %p')}")

                # Find free slots
                current_time = day_start
                day_free_slots = set()

                #logging.info(f"Checking time slots for {current_date.strftime('%A %d %B')}:")

                while current_time + timedelta(hours=1) <= day_end:
                    slot_end = current_time + timedelta(hours=1)
                    is_free = True

                    for busy_start, busy_end, summary in busy_times:
                        if (busy_start < slot_end and busy_end > current_time):
                            #logging.info(
                            #    f"  BUSY : {current_time.strftime('%I:%M %p')} - {slot_end.strftime('%I:%M %p')} "
                            #    f"(Blocked by: {summary})")
                            is_free = False
                            break

                    if is_free:
                        day_free_slots.add(current_time)
                        #logging.info(f"  FREE : {current_time.strftime('%I:%M %p')} - {slot_end.strftime('%I:%M %p')}")

                    current_time += timedelta(hours=1)

                if day_free_slots:
                    free_slots[current_date] = day_free_slots
                    #logging.info(f"Final free slots for {current_date.strftime('%A %d %B')}:")
                    #for slot in sorted(day_free_slots):
                        #logging.info(
                        #    f"  {slot.strftime('%I:%M %p')} - {(slot + timedelta(hours=1)).strftime('%I:%M %p')}")

            elapsed_time = time.time() - start_time
            logging.info(f"Calendar loaded in {elapsed_time:.2f} seconds: {url}")
            return free_slots

        except Exception as e:
            elapsed_time = time.time() - start_time
            logging.error(f"Error fetching calendar ({elapsed_time:.2f} seconds): {url} - {e}")
            logging.error(f"Exception details:", exc_info=True)
            return {}

    def find_common_free_slots(self):
        """Finds common free slots across multiple calendars."""
        start_time = time.time()
        all_free_slots = []

        for url in self.calendar_urls:
            free_slots = self.fetch_and_parse_calendar(url)
            all_free_slots.append(free_slots)

        common_free_slots = {}
        all_dates = set()
        for calendar_slots in all_free_slots:
            all_dates.update(calendar_slots.keys())

        now = datetime.now(self.local_tz)
        current_date = now.date()

        for slot_date in all_dates:
            # Skip current day if include_current_day is False
            if not self.include_current_day and slot_date == current_date:
                continue

            date_slots = []
            for calendar_slots in all_free_slots:
                if slot_date in calendar_slots:
                    date_slots.append(calendar_slots[slot_date])
                else:
                    break

            if len(date_slots) == len(self.calendar_urls):
                common_slots = set.intersection(*date_slots)
                if common_slots:
                    # For current day, filter out time slots that have already passed
                    if slot_date == current_date:
                        current_time = now.replace(minute=0, second=0, microsecond=0)
                        common_slots = {slot for slot in common_slots if slot > current_time}

                    if common_slots:  # Only add if there are slots remaining
                        common_free_slots[slot_date] = sorted(list(common_slots))

        elapsed_time = time.time() - start_time
        logging.info(f"Total calendar update completed in {elapsed_time:.2f} seconds")
        return common_free_slots

    def format_free_slots(self, free_slots):
        """Formats the free slots into the requested output format."""
        def ordinal(n):
            return f"{n}{'tsnrhtdd'[((n // 10 % 10 != 1) * (n % 10 < 4) * n % 10)::4]}"

        formatted_output = [self.custom_text]
        for date, slots in sorted(free_slots.items()):
            day_str = date.strftime("%a") + f" {ordinal(date.day)}" + date.strftime(" %b")
            slot_strs = [slot.strftime("%I%p").lstrip("0").lower() for slot in slots]
            formatted_output.append(f"{day_str}: {', '.join(slot_strs)}")
        return "\n".join(formatted_output)

    def toggle_weekends(self):
        """Toggle weekend exclusion and update free slots"""
        self.exclude_weekends = not self.exclude_weekends
        self.update_free_slots()

    def paste_free_slots(self):
        """Copy free slots to clipboard and simulate paste."""
        logging.info("Paste free slots triggered")
        try:
            if self.cached_free_slots:
                formatted_text = self.format_free_slots(self.cached_free_slots)

                # Store original clipboard content
                self.original_clipboard = pyperclip.paste()

                # Copy our text
                pyperclip.copy(formatted_text)
                logging.info("Text copied to clipboard")

                # Small delay before paste
                time.sleep(0.1)

                # Simulate paste
                if platform.system() == 'Darwin':  # macOS
                    pyautogui.hotkey('command', 'v')
                else:  # Windows/Linux
                    pyautogui.hotkey('ctrl', 'v')

                logging.info("Paste command sent")

                # Restore original clipboard after a delay
                threading.Timer(0.5, self.restore_clipboard).start()

                logging.info("Paste operation completed")
            else:
                logging.warning("No cached free slots available")
        except Exception as e:
            logging.error(f"Error pasting free slots: {e}", exc_info=True)

    def update_free_slots(self):
        """Update cached free slots."""
        if hasattr(self, '_update_in_progress') and self._update_in_progress:
            logging.info("Calendar update already in progress, skipping new update request")
            return

        def update_task():
            try:
                self._update_in_progress = True
                logging.info("Starting calendar update...")
                self.cached_free_slots = self.find_common_free_slots()
                self.save_cache()
                logging.info("Calendar update completed successfully")
            except Exception as e:
                logging.error(f"Error updating free slots: {e}")
            finally:
                self._update_in_progress = False

        threading.Thread(target=update_task, daemon=True).start()

    def update_loop(self):
        """Background loop to update free slots periodically."""
        while True:
            self.update_free_slots()
            time.sleep(self.update_interval)

    def load_icon(self):
        """Load icon from embedded resource or create default"""
        try:
            # Try to load embedded icon
            icon_path = get_resource_path('icon.png')
            icon_image = Image.open(icon_path)
            logging.info("Custom icon loaded successfully")
            return icon_image
        except Exception as e:
            logging.warning(f"Could not load icon: {e}, using default")
            # Create a default icon (your existing default icon code)
            image = Image.new('RGB', (64, 64), 'white')
            pixels = image.load()
            # Draw border
            for i in range(64):
                for j in range(64):
                    if i < 2 or i > 61 or j < 2 or j > 61:
                        pixels[i, j] = (0, 0, 0)
                    elif j < 15:
                        pixels[i, j] = (200, 200, 200)
            return image

    def create_icon(self):
        """Return the loaded icon"""
        return self.icon_image


    def cleanup(self):
        """Cleanup before exit"""
        logging.info("FreeTime application shutting down...")
        try:
            keyboard.unhook_all()
            logging.info("Keyboard hooks removed")

            if self.original_clipboard is not None:
                pyperclip.copy(self.original_clipboard)
        except Exception as e:
            logging.error(f"Error during cleanup: {e}")

        if self.icon:
            self.icon.stop()
        self.root.quit()

    def show_about(self):
        """Show the About window"""
        try:
            if self.about_window is None or not tk.Toplevel.winfo_exists(self.about_window.window):
                self.root.deiconify()  # Ensure root window exists
                self.about_window = AboutWindow(self.root, self.icon_image)
                self.root.withdraw()  # Hide root window again
            else:
                self.about_window.window.lift()
                self.about_window.window.focus_force()
        except Exception as e:
            logging.error(f"Error showing About window: {e}")

    def run(self):
        """Run the application."""
        try:
            # Create and start update thread
            update_thread = threading.Thread(target=self.update_loop, daemon=True)
            update_thread.start()

            # Setup hotkey
            self.setup_hotkey()

            # Create system tray icon
            self.icon = pystray.Icon(
                "Calendar",
                self.create_icon(),
                "FreeTime",
                menu=pystray.Menu(
                    pystray.MenuItem("About", lambda: self.root.after(0, self.show_about)),  # Add this line
                    pystray.MenuItem("Settings", lambda: self.root.after(0, self.show_settings)),
                    pystray.MenuItem("Update Now", lambda: self.update_free_slots()),
                    pystray.MenuItem("Exit", lambda: self.cleanup())
                )
            )

            logging.info("System tray icon initialized")

            # Check if calendar URLs exist, if not show settings
            if not self.calendar_urls:
                self.root.after(1000, self.show_settings)

            # Run the icon in a separate thread
            icon_thread = threading.Thread(target=self.icon.run, daemon=True)
            icon_thread.start()

            # Start the tkinter main loop
            self.root.mainloop()

        except Exception as e:
            logging.error(f"Fatal error in main loop: {e}")
            raise

if __name__ == "__main__":
    try:
        logging.info("=" * 50)
        logging.info("FreeTime Application Session Start")
        logging.info(f"Python version: {sys.version}")
        logging.info(f"Operating System: {platform.system()} {platform.version()}")

        app = CalendarApp()
        app.run()
    except Exception as e:
        logging.error(f"Application error: {e}")
        raise