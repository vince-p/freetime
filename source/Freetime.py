import requests
from datetime import datetime, timedelta
import pytz
import pyperclip
from typing import List, Dict, Set
from datetime import date
import pystray
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
from PIL import Image, ImageTk
from icalendar import Calendar
from typing import List, Dict, Set, Tuple


version_string = "0.99"
# Changelog
#1

# v0.99 8/5/25
# Working well across platforms

# v091 1/5/25
# Fix logging code so that it works across platforms

# v09 25/4/25
# fixed an error where recurring events created before daylight saving change are processed incorrectly after daylight savings.

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
            lines = []
            if os.path.exists(self.baseFilename):
                try:
                    # Explicitly use utf-8 encoding and ignore errors on read
                    with open(self.baseFilename, 'r', encoding='utf-8', errors='ignore') as f:
                        lines = f.readlines()
                except Exception as e:
                    # Log an error if reading fails, but continue
                    logging.error(f"Error reading log file {self.baseFilename}: {e}", exc_info=True)
                    lines = [] # Start fresh if read fails

            # Add new log entry at the beginning
            new_log_line = self.format(record) + '\n'
            lines.insert(0, new_log_line)

            # Keep only the most recent MAX_ENTRIES
            lines = lines[:self.MAX_ENTRIES]

            try:
                # Explicitly use utf-8 encoding for writing
                with open(self.baseFilename, 'w', encoding='utf-8') as f:
                    f.writelines(lines)
            except Exception as e:
                 logging.error(f"Error writing log file {self.baseFilename}: {e}", exc_info=True)
                 # If writing fails, we might lose logs, but the app shouldn't crash

        except Exception:
            self.handleError(record) # Use standard error handling if something else goes wrong

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

        system = platform.system()

        # --- Platform-specific initial window state ---
        if system == 'Windows':
            # On Windows, withdraw immediately to prevent flash
            self.window.withdraw()
            # Setting an initial geometry can sometimes help with early layout calculations on Windows
            self.window.geometry("300x230")
        else: # For macOS and other systems
            # On macOS (and others), do NOT withdraw initially.
            # The window will appear immediately, and we will center it shortly after.
            pass
        # --- End Platform-specific initial window state ---


        # Make window non-resizable (same for all platforms)
        self.window.resizable(False, False)

        # Set window icon (same for all platforms)
        try:
            icon_photo = ImageTk.PhotoImage(icon_image)
            self.window.iconphoto(True, icon_photo)
            self.icon_photo = icon_photo  # Keep a reference
        except Exception as e:
            logging.error(f"Error setting window icon: {e}")

        # --- Build Content (same for all platforms) ---
        # Create main frame with padding
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(expand=True, fill='both')

        # Create inner frame for content
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(expand=True)

        # Display logo
        try:
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
        # --- End Build Content ---

        # --- Platform-specific Centering Logic ---
        if system == 'Windows':
            # On Windows, calculate size and position while withdrawn, then deiconify
            try:
                # Force Tkinter to calculate the required size based on packed content
                self.window.update()
                self.window.update_idletasks()

                # Get the window's actual size based on content
                content_width = self.window.winfo_width()
                content_height = self.window.winfo_height()

                # Increase dimensions by 20%
                new_width = int(content_width * 1.2)
                new_height = int(content_height * 1.2)

                # Get screen dimensions
                screen_width = self.window.winfo_screenwidth()
                screen_height = self.window.winfo_screenheight()

                # Calculate position x, y coordinates for centering the *new* size
                x = (screen_width - new_width) // 2
                y = (screen_height - new_height) // 2

                logging.debug(f"Windows Centering: content_size={content_width}x{content_height}, new_size={new_width}x{new_height}, screen={screen_width}x{screen_height}, pos={x}+{y}")

                # Set the window's new size and position while it's still withdrawn
                self.window.geometry(f"{new_width}x{new_height}+{x}+{y}")

            except Exception as e:
                 logging.error(f"Error during Windows centering calculation: {e}", exc_info=True)

            # Make the window visible *after* positioning
            self.window.deiconify()

        else: # macOS and others
            # On macOS, schedule delayed centering
            # The window is already visible at this point because we didn't withdraw
            self.window.after(50, self._center_window_mac) # Use a small delay like 50ms

        # --- End Platform-specific Centering Logic ---


        # Ensure window is on top and gets focus after being shown/positioned
        # This happens after deiconify on Windows or after scheduling on macOS
        self.window.lift()
        self.window.attributes('-topmost', True)
        # Call update() again here for robustness after state changes
        self.window.update()
        self.window.attributes('-topmost', False)
        self.window.focus_force()


    def _center_window_mac(self):
        """Calculates and sets the window's position to center it (for macOS/other)."""
        try:
            # Force geometry calculation
            self.window.update_idletasks()
            # Use update() here as well for reliability before querying dimensions
            self.window.update()

            # Get the window's actual size based on content
            content_width = self.window.winfo_width()
            content_height = self.window.winfo_height()

            # Increase dimensions by 20% (same logic as Windows)
            new_width = int(content_width * 1.2)
            new_height = int(content_height * 1.2)

            # Get screen dimensions
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()

            # Calculate position x, y coordinates for centering the *new* size
            x = (screen_width - new_width) // 2
            y = (screen_height - new_height) // 2

            logging.debug(f"macOS Centering: content_size={content_width}x{content_height}, new_size={new_width}x{new_height}, screen={screen_width}x{screen_height}, pos={x}+{y}")

            # Set the window's new size and position
            self.window.geometry(f"{new_width}x{new_height}+{x}+{y}")

        except Exception as e:
             logging.error(f"Error during macOS window centering: {e}", exc_info=True)

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

        # Startup checkbox
        self.startup_var = tk.BooleanVar(value=self.check_startup())
        startup_cb = ttk.Checkbutton(app_frame, text="Run at startup",
                                     variable=self.startup_var,
                                     command=self.toggle_startup)
        startup_cb.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)

        current_row += 1

        # Add an empty spacer row with weight to push the bottom frame down
        # This is the key change - add a spacer frame that can expand
        spacer_frame = ttk.Frame(main_frame)
        spacer_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.rowconfigure(current_row, weight=1)  # Give this row weight to expand
        current_row += 1

        # Now create the bottom frame in the next row
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=current_row, column=0, pady=10, sticky=(tk.W, tk.E, tk.S))
        bottom_frame.columnconfigure(1, weight=1)  # Make middle column expand

        # Log File Link (left aligned)
        log_label = ttk.Label(bottom_frame, text="View log", cursor="hand2")
        log_label.grid(row=0, column=0, sticky=tk.W)
        log_label.bind("<Button-1>", self.open_log_file)

        # Get the font used by other ttk elements
        default_font = ttk.Style().lookup('TLabel', 'font')

        # Ensure we have a tk_font.Font object, fallback to 'TkDefaultFont' if needed
        try:
            name = str(default_font) if default_font else "TkDefaultFont"
            default_font_obj = tk_font.nametofont(name)
        except Exception:
            default_font_obj = tk_font.nametofont("TkDefaultFont")

        # Now we can safely use cget
        link_font = tk_font.Font(
            family=default_font_obj.cget("family"),
            size=max(default_font_obj.cget("size") - 1, 1),  # avoid size < 1
            weight=default_font_obj.cget("weight")
        )
        link_font.configure(underline=True)
        log_label.configure(font=link_font)

        # Status Label (center aligned)
        self.status_label = ttk.Label(bottom_frame, text="", foreground="green")
        self.status_label.grid(row=0, column=1)

        # Save Button (right aligned)
        save_button = ttk.Button(bottom_frame, text="Save", command=self.save_settings)
        save_button.grid(row=0, column=2, sticky=tk.E)

        # Ensure the window is large enough to show all content
        self.window.update_idletasks()
        min_height = 740  # Set minimum height
        current_height = self.window.winfo_height()
        if current_height < min_height:
            self.window.geometry(f"{self.window.winfo_width()}x{min_height}")

    def save_settings(self):
        """Save settings, with special handling for the trigger pattern on macOS"""
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
            trigger_changed = new_trigger != old_trigger

            # Handle trigger pattern change
            if trigger_changed:
                logging.info(f"Trigger changing from '{old_trigger}' to '{new_trigger}'")

                # Update the main trigger pattern
                self.app.trigger_pattern = new_trigger

                # Update the trigger patterns list
                if hasattr(self.app, 'trigger_patterns'):
                    # Add the new trigger to the beginning of the list
                    self.app.trigger_patterns.insert(0, new_trigger.lower())
                    # Keep the old trigger temporarily to allow smooth transition
                    if old_trigger.lower() in self.app.trigger_patterns:
                        self.app.trigger_patterns.remove(old_trigger.lower())
                    # Keep only the last 2 patterns maximum
                    self.app.trigger_patterns = self.app.trigger_patterns[:2]
                else:
                    self.app.trigger_patterns = [new_trigger.lower()]

                logging.info(f"Updated trigger patterns list: {self.app.trigger_patterns}")

                # No need to restart the listener on macOS - it will use the updated patterns list
                # For Windows, we need to update the hotkey
                if platform.system() == "Windows":
                    self.app.setup_hotkey()

            # Update all other settings
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
            logging.error(error_msg, exc_info=True)
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

    def ensure_save_button_visible(self):
        """Force the save button to be visible and properly rendered"""
        try:
            # Find all frames in the window
            all_frames = self._find_all_widgets(self.window, ttk.Frame)
            bottom_frame = None

            # Find the bottom frame
            for frame in all_frames:
                grid_info = frame.grid_info()
                if grid_info.get('row') == len(all_frames) - 1:  # Last row
                    bottom_frame = frame
                    break

            if bottom_frame:
                # Recreate save button if needed
                save_button = ttk.Button(bottom_frame, text="Save", command=self.save_settings)
                save_button.grid(row=0, column=2, sticky=tk.E)
                bottom_frame.update()
                self.window.update_idletasks()
        except Exception as e:
            logging.error(f"Error ensuring save button is visible: {e}")

    def _find_all_widgets(self, parent, widget_type):
        """Helper to find all widgets of a specific type"""
        result = []
        for child in parent.winfo_children():
            if isinstance(child, widget_type):
                result.append(child)
            result.extend(self._find_all_widgets(child, widget_type))
        return result


class CalendarApp:
    def __init__(self):
        self.settings_file = SETTINGS_FILE
        self.cache_file = APP_DIR / 'calendar_cache.json'

        # Load icon first
        self.icon_image = self.load_icon()

        # Log startup
        logging.debug(f"App directory: {APP_DIR}")
        logging.info(f"Settings file: {self.settings_file}")
        logging.info(f"Cache file: {self.cache_file}")

        # Initialize threading objects
        self.paste_lock = threading.Lock()
        self.is_pasting = False

        # Initialize trigger patterns as a list - the active one will be first
        self.trigger_patterns = []

        self.load_settings()

        # Ensure trigger pattern is in the patterns list
        if hasattr(self, 'trigger_pattern'):
            self.trigger_patterns = [self.trigger_pattern.lower()]

        self.cached_free_slots = None
        self.load_cache()
        self.settings_window = None

        # Initialize the root window
        self.root = tk.Tk()
        self.root.withdraw()
        self.icon = None

        self.original_clipboard = None
        self.about_window = None

    def setup_hotkey(self):
        """
        Set up cross-platform text pattern detection:
        - Use 'keyboard' on Windows
        - Use 'pynput' on macOS
        """
        system = platform.system()
        self.char_buffer = ""
        self.buffer_max_size = 20

        # Ensure we have the paste lock
        if not hasattr(self, 'paste_lock'):
            self.paste_lock = threading.Lock()

        # Ensure we have trigger patterns list
        if not hasattr(self, 'trigger_patterns') or not self.trigger_patterns:
            self.trigger_patterns = [self.trigger_pattern.lower()]

        # Only start a new listener if we don't have one already
        if system == "Darwin" and (not hasattr(self, '_keyboard_listener') or not self._keyboard_listener):
            from pynput import keyboard as pynput_keyboard

            logging.info(f"Setting up hotkey monitor for macOS with trigger patterns: {self.trigger_patterns}")

            def on_press(key):
                # Only process keys if we're not currently pasting
                if hasattr(self, 'is_pasting') and self.is_pasting:
                    return

                try:
                    if hasattr(key, "char") and key.char is not None:
                        c = key.char
                        self.char_buffer += c
                        if len(self.char_buffer) > self.buffer_max_size:
                            self.char_buffer = self.char_buffer[-self.buffer_max_size:]
                        logging.debug(f"Buffer now: {repr(self.char_buffer)}")

                        # Check against all trigger patterns, with the active one first
                        trigger_found = False
                        for trigger in self.trigger_patterns:
                            if trigger in self.char_buffer:
                                logging.debug(f"Trigger '{trigger}' detected")
                                trigger_found = True
                                break

                        if trigger_found:
                            # Clear buffer immediately
                            self.char_buffer = ""
                            # Schedule trigger_paste with a small delay
                            if hasattr(self, 'root') and self.root:
                                self.root.after(10, self.trigger_paste)
                            else:
                                threading.Timer(0.01, self.trigger_paste).start()
                    elif key == pynput_keyboard.Key.space:
                        self.char_buffer += " "
                    elif key == pynput_keyboard.Key.enter:
                        self.char_buffer = ""
                    elif key == pynput_keyboard.Key.backspace:
                        self.char_buffer = self.char_buffer[:-1]
                except Exception as e:
                    logging.error(f"Error in on_press: {e}")

            # Create and start listener
            try:
                listener = pynput_keyboard.Listener(on_press=on_press, daemon=True)
                listener.start()
                self._keyboard_listener = listener
                logging.info("Keyboard listener started successfully for macOS")
            except Exception as e:
                logging.error(f"Failed to start keyboard listener: {e}")
                self._keyboard_listener = None

        elif system == "Windows":
            # Windows setup (unchanged)
            try:
                import keyboard
                keyboard.unhook_all()

                def on_key(event):
                    if self.is_pasting:
                        return
                    if self.settings_window is not None and tk.Toplevel.winfo_exists(self.settings_window.window):
                        active_widget = self.settings_window.window.focus_get()
                        if active_widget and isinstance(active_widget, (ttk.Entry, tk.Entry, ttk.Combobox)):
                            return

                    if event.event_type == keyboard.KEY_DOWN and hasattr(event, 'name'):
                        if event.name == 'space':
                            self.char_buffer += ' '
                        elif event.name == 'enter':
                            self.char_buffer = ""
                        elif event.name == 'backspace':
                            self.char_buffer = self.char_buffer[:-1]
                        elif len(event.name) == 1:
                            self.char_buffer += event.name
                        if len(self.char_buffer) > self.buffer_max_size:
                            self.char_buffer = self.char_buffer[-self.buffer_max_size:]
                        if self.trigger_pattern in self.char_buffer:
                            for _ in range(len(self.trigger_pattern)):
                                keyboard.send('backspace')
                            self.char_buffer = ""
                            self.trigger_paste()

                keyboard.hook(on_key)
                self._keyboard_listener = None  # not used for keyboard lib
                logging.info(f"Keyboard hotkey hooked using 'keyboard' for Windows")

            except Exception as e:
                logging.error(f"Failed to set up hotkey for Windows: {e}")

    def _setup_mac_listener(self):
        """
        Set up the keyboard listener for macOS in a safe way
        """
        try:
            from pynput import keyboard as pynput_keyboard

            # Make sure we're using the most current trigger phrase
            trigger = self.trigger_pattern.lower()
            logging.info(f"Setting up hotkey monitor for macOS, trigger: {trigger!r}")

            def on_press(key):
                # Skip processing during paste operations
                if hasattr(self, 'is_pasting') and self.is_pasting:
                    return

                try:
                    if hasattr(key, "char") and key.char is not None:
                        c = key.char
                        self.char_buffer += c
                        if len(self.char_buffer) > self.buffer_max_size:
                            self.char_buffer = self.char_buffer[-self.buffer_max_size:]
                        logging.debug(f"Buffer now: {repr(self.char_buffer)}")

                        # Always check against the current trigger pattern
                        current_trigger = self.trigger_pattern.lower()
                        if current_trigger in self.char_buffer:
                            logging.debug(f"Trigger '{current_trigger}' detected")
                            # Clear buffer immediately
                            self.char_buffer = ""
                            # Schedule with a delay to make it thread-safe
                            if hasattr(self, 'root') and self.root:
                                self.root.after(10, self.trigger_paste)
                            else:
                                # Fallback if root is not available
                                threading.Timer(0.01, self.trigger_paste).start()
                    elif key == pynput_keyboard.Key.space:
                        self.char_buffer += " "
                    elif key == pynput_keyboard.Key.enter:
                        self.char_buffer = ""
                    elif key == pynput_keyboard.Key.backspace:
                        self.char_buffer = self.char_buffer[:-1]
                except Exception as e:
                    logging.error(f"Error in on_press: {e}")

            # Create a new listener with daemon=True to ensure it doesn't block application exit
            listener = pynput_keyboard.Listener(
                on_press=on_press,
                daemon=True
            )

            # Start the listener
            listener.start()
            self._keyboard_listener = listener
            logging.info("Keyboard listener started successfully for macOS")

        except Exception as e:
            logging.error(f"Failed to set up macOS keyboard listener: {e}")
            self._keyboard_listener = None

    def _delayed_setup_mac_listener(self):
        """
        Set up the Mac keyboard listener after a delay
        """
        logging.info("Running delayed listener setup")
        self._recently_stopped_listener = False
        self._setup_mac_listener()

    def _stop_all_keyboard_listeners(self):
        """
        Thoroughly clean up any existing keyboard listeners to prevent conflicts
        """
        # Stop our known listener if it exists
        if hasattr(self, '_keyboard_listener') and self._keyboard_listener:
            try:
                logging.info("Stopping existing keyboard listener")
                self._keyboard_listener.stop()
                self._keyboard_listener = None
            except Exception as e:
                logging.error(f"Error stopping keyboard listener: {e}")
                self._keyboard_listener = None

        # For Windows, unhook all existing keyboard hooks
        try:
            if platform.system() == "Windows":
                import keyboard
                keyboard.unhook_all()
                logging.info("Unhooked all keyboard hooks on Windows")
        except Exception as e:
            logging.error(f"Error clearing Windows keyboard hooks: {e}")

        # Give a small delay to ensure cleanup completes
        time.sleep(0.1)

    def trigger_paste(self):
        """Method to handle trigger and paste operation"""
        # Make sure we have a paste lock
        if not hasattr(self, 'paste_lock'):
            self.paste_lock = threading.Lock()

        # Use a lock to prevent multiple paste operations
        if not self.paste_lock.acquire(blocking=False):
            logging.debug("Another paste operation in progress, ignoring trigger")
            return

        try:
            # Set pasting flag to ignore keyboard events during paste
            self.is_pasting = True
            logging.debug("Starting paste operation with lock acquired")

            if platform.system() == 'Darwin':
                # First, delete the trigger text
                for _ in range(len(self.trigger_pattern)):
                    pyautogui.press('backspace')

                # Small delay after backspace
                time.sleep(0.05)

                # Now perform the clipboard-based paste operation
                if self.cached_free_slots:
                    # Format the text
                    formatted_text = self.format_free_slots(self.cached_free_slots)

                    # Store original clipboard content
                    try:
                        original_clip = pyperclip.paste()
                        self.original_clipboard = original_clip
                    except Exception as e:
                        logging.error(f"Error saving clipboard: {e}")
                        self.original_clipboard = None

                    # Copy our text to clipboard
                    pyperclip.copy(formatted_text)
                    logging.debug("Text copied to clipboard")

                    # Small delay to ensure clipboard is updated
                    time.sleep(0.1)

                    # Paste using Command+V
                    try:
                        # Use proper key sequence with small delays
                        pyautogui.keyDown('command')
                        time.sleep(0.05)
                        pyautogui.press('v')
                        time.sleep(0.05)
                        pyautogui.keyUp('command')
                        logging.debug("Paste command sent")
                    except Exception as e:
                        logging.error(f"Error sending paste command: {e}")

                    # Restore original clipboard after delay
                    def restore_clip():
                        if self.original_clipboard is not None:
                            try:
                                pyperclip.copy(self.original_clipboard)
                                self.original_clipboard = None
                                logging.debug("Original clipboard restored")
                            except Exception as e:
                                logging.error(f"Error restoring clipboard: {e}")

                    # Use slightly longer delay for clipboard restoration
                    threading.Timer(1.0, restore_clip).start()
                else:
                    logging.warning("No cached free slots available")
            else:
                # For Windows/Linux, use the original method
                self.paste_free_slots()

        except Exception as e:
            logging.error(f"Error in trigger_paste: {e}", exc_info=True)
        finally:
            # Reset paste flag after a delay
            def reset_pasting():
                self.is_pasting = False
                try:
                    self.paste_lock.release()
                    logging.debug("Paste lock released")
                except Exception as e:
                    logging.error(f"Error releasing paste lock: {e}")

            threading.Timer(0.5, reset_pasting).start()

    def debug_keyboard_state(self):
        """Debug helper to print current keyboard state"""
        logging.debug("=== Keyboard State Debug ===")
        logging.debug(f"Hotkey keys: {self.hotkey_keys}")
        logging.debug(f"Current keys: {self.current_keys}")
        #logging.info(f"Time since last trigger: {time.time() - self.last_trigger_time}")
        logging.debug("==========================")

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
            logging.debug(f"Attempting to load settings from: {self.settings_file}")

            if self.settings_file.exists():
                with open(self.settings_file, 'r') as f:
                    file_content = f.read()
                    logging.debug(f"Settings file content: {file_content}")

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
                            logging.debug(f"Loaded setting {key}: {value}")

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
                'custom_text': self.custom_text
            }

            # Log what we're about to save
            logging.info(f"Saving settings to: {self.settings_file}")
            logging.info(f"Settings data: {settings}")

            # Create directory if it doesn't exist
            self.settings_file.parent.mkdir(parents=True, exist_ok=True)

            with open(self.settings_file, 'w') as f:
                json_data = json.dumps(settings, indent=2)
                f.write(json_data)

            logging.debug("Settings saved successfully")
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
        """Show the settings window with improved window handling for compiled app"""
        try:
            # Force any old window to be destroyed first
            if hasattr(self, 'settings_window') and self.settings_window is not None:
                try:
                    if hasattr(self.settings_window, 'window') and self.settings_window.window is not None:
                        if tk.Toplevel.winfo_exists(self.settings_window.window):
                            self.settings_window.window.destroy()
                except Exception as e:
                    logging.debug(f"Error cleaning up old settings window: {e}")

            # Create a fresh settings window
            self.settings_window = SettingsWindow(self, self.root)

            # Ensure it's visible and on top
            self.settings_window.window.lift()
            self.settings_window.window.attributes('-topmost', True)
            self.settings_window.window.update()
            self.settings_window.window.attributes('-topmost', False)
            self.settings_window.window.focus_force()

            # Log for debugging
            logging.debug("Settings window created and shown")

            # Validate that the Save button exists and is properly configured
            save_buttons = [w for w in self.settings_window.window.winfo_children()
                            if hasattr(w, 'winfo_class') and w.winfo_class() == 'TButton'
                            and w.cget('text') == 'Save']
            #logging.debug(f"Found {len(save_buttons)} Save buttons in the window")

        except Exception as e:
            logging.error(f"Error showing settings window: {e}", exc_info=True)

    def is_weekend(self, date_obj):
        """Check if the given date is a weekend (Saturday=5 or Sunday=6)"""
        return date_obj.weekday() >= 5

    # This is a helper function. The only purpose of this is to allow easy debugging of busy times.
    def format_busy_log(busy_dict: Dict[date, List[Tuple[datetime, datetime, str]]], start_date: date, lookahead_days: int, local_tz) -> str:
        """Formats the busy dictionary for readable debug output. (Static Method)"""
        log_lines = ["--- Busy Event Log ---"]
        end_date = start_date + timedelta(days=lookahead_days - 1) # Calculate end date based on lookahead

        # Get all relevant dates within the lookahead window
        relevant_dates = sorted([d for d in busy_dict.keys() if start_date <= d <= end_date])

        # Create a set of dates we actually have data for within the range
        dates_with_data = set(relevant_dates)

        # Iterate through each day in the lookahead range
        for i in range(lookahead_days):
            current_day = start_date + timedelta(days=i)
            day_str = current_day.strftime("%Y-%m-%d %A")
            log_lines.append(f"=== {day_str} ===")

            if current_day in dates_with_data:
                # Sort events by start time for the current day
                sorted_events = sorted(busy_dict[current_day], key=lambda x: x[0])
                if sorted_events:
                    for start_time, end_time, summary in sorted_events:
                        # Format start and end times
                        start_f = start_time.astimezone(local_tz).strftime("%H:%M")
                        end_f = end_time.astimezone(local_tz).strftime("%H:%M")

                        # If the event spans across midnight into the next day, show the date too
                        if end_time.astimezone(local_tz).date() > start_time.astimezone(local_tz).date():
                             end_f = end_time.astimezone(local_tz).strftime("%Y-%m-%d %H:%M")

                        log_lines.append(f"  - Busy: {start_f} to {end_f} | Event: {summary}")
                else:
                    log_lines.append("  No busy events recorded for this day.")
            else:
                log_lines.append("  No busy events recorded for this day.")

        log_lines.append("--- End Busy Event Log ---")
        return "\n".join(log_lines)

    def fetch_and_parse_calendar(self, url: str) -> Dict[date, Set[datetime]]:
        """
        Parse *url* and return {date: set(tz-aware datetime start-times)}
        representing one-hour BUSY slots.  Handles RRULE/RDATE/EXDATE
        across daylight-saving changes and eliminates UNTIL/DTSTART
        timezone conflicts.
        """
        t0 = time.time()
        try:
            gcal = Calendar.from_ical(requests.get(url, timeout=15).text)

            today     = datetime.now(self.local_tz).date()
            win_end   = today + timedelta(days=self.lookahead_days)
            busy: Dict[date, List[Tuple[datetime, datetime, str]]] = {}

            def _as_list(prop):
                if prop is None:
                    return []
                return prop if isinstance(prop, list) else [prop]

            for comp in gcal.walk():
                if comp.name != "VEVENT":
                    continue
                if comp.get("status", "").upper() == "CANCELLED":
                    continue          # cancelled master

                summary  = str(comp.get("summary", "")) or "No title"

                dt_start = comp.decoded("dtstart")
                dt_end   = comp.decoded("dtend")

                # all-day dates → midnight datetimes
                if isinstance(dt_start, date) and not isinstance(dt_start, datetime):
                    dt_start = datetime.combine(dt_start, datetime.min.time())
                if isinstance(dt_end,   date) and not isinstance(dt_end,   datetime):
                    dt_end   = datetime.combine(dt_end,   datetime.min.time())

                # attach zone if absent
                if dt_start.tzinfo is None:
                    dt_start = self.local_tz.localize(dt_start)
                if dt_end.tzinfo is None:
                    dt_end   = self.local_tz.localize(dt_end)

                dt_start = dt_start.astimezone(self.local_tz)
                dt_end   = dt_end  .astimezone(self.local_tz)
                duration = dt_end - dt_start

                is_all_day = isinstance(comp.get("dtstart").dt, date) and \
                             not isinstance(comp.get("dtstart").dt, datetime)
                multi_day  = (dt_end.date() - dt_start.date()).days > 0
                if (is_all_day or multi_day) and self.ignore_all_day_events:
                    continue

                # ───────────── overrides (RECURRENCE-ID) ─────────────
                if comp.get("recurrence-id"):
                    rid = comp.decoded("recurrence-id")
                    if rid.tzinfo is None:
                        rid = self.local_tz.localize(rid)
                    rid = rid.astimezone(self.local_tz)

                    # remove original instance
                    for blk in busy.get(rid.date(), [])[:]:
                        if abs((blk[0] - rid).total_seconds()) < 1:
                            busy[rid.date()].remove(blk)
                            break
                    busy.setdefault(dt_start.date(), []).append((dt_start, dt_end, summary))
                    continue

                rrule_prop = comp.get("rrule")

                # ───────────── recurring events ─────────────
                if rrule_prop:
                    from dateutil import rrule as dtr
                    import pytz, re

                    rrule_txt = rrule_prop.to_ical().decode("utf-8")

                    # fix local UNTIL → UTC
                    m = re.search(r"UNTIL=(\d{8}T\d{6})(Z?)", rrule_txt)
                    if m and not m.group(2):
                        until_local = datetime.strptime(m.group(1), "%Y%m%dT%H%M%S")
                        until_local = self.local_tz.localize(until_local)
                        until_utc   = until_local.astimezone(pytz.utc)
                        rrule_txt   = (rrule_txt[:m.start(1)] +
                                       until_utc.strftime("%Y%m%dT%H%M%SZ") +
                                       rrule_txt[m.end(1):])

                    rset = dtr.rruleset()
                    # ignoretz=True ⇒ every value (DTSTART/UNTIL/BYxxx) treated naive
                    rset.rrule(dtr.rrulestr(rrule_txt,
                                            dtstart=dt_start.replace(tzinfo=None),
                                            ignoretz=True))

                    for rdate in _as_list(comp.get("rdate")):
                        for rdt in rdate.dts:
                            rset.rdate(rdt.dt.astimezone(self.local_tz).replace(tzinfo=None))
                    for exdate in _as_list(comp.get("exdate")):
                        for exdt in exdate.dts:
                            rset.exdate(exdt.dt.astimezone(self.local_tz).replace(tzinfo=None))

                    win_start_nv = datetime.combine(today, datetime.min.time())
                    win_end_nv   = datetime.combine(win_end, datetime.max.time())

                    for occ_nv in rset.between(win_start_nv, win_end_nv, inc=True):
                        occ_start = self.local_tz.localize(occ_nv, is_dst=None)
                        occ_end   = occ_start + duration
                        busy.setdefault(occ_start.date(), []).append(
                            (occ_start, occ_end, summary))
                    continue

                # ───────────── single / multi-day non-recurring ─────────────
                if multi_day:
                    ptr = dt_start.date()
                    while ptr <= dt_end.date():
                        blk_s = dt_start if ptr == dt_start.date() else \
                                self.local_tz.localize(datetime.combine(ptr, datetime.min.time()))
                        blk_e = dt_end   if ptr == dt_end.date()   else \
                                self.local_tz.localize(datetime.combine(ptr, datetime.max.time()))
                        busy.setdefault(ptr, []).append((blk_s, blk_e, summary))
                        ptr += timedelta(days=1)
                else:
                    busy.setdefault(dt_start.date(), []).append((dt_start, dt_end, summary))

            # Format and log the busy dictionary content for the relevant range
            formatted_log = CalendarApp.format_busy_log(busy, today, self.lookahead_days, self.local_tz)
            logging.debug(f"Processed busy blocks for URL {url}:\n{formatted_log}")

            # ───────────── derive FREE one-hour slots ─────────────
            free: Dict[date, Set[datetime]] = {}
            for i in range(self.lookahead_days):
                d = today + timedelta(days=i)
                if self.exclude_weekends and self.is_weekend(d):
                    continue
                day_s = self.local_tz.localize(
                    datetime.combine(d, datetime.min.time())).replace(hour=self.start_of_day)
                day_e = self.local_tz.localize(
                    datetime.combine(d, datetime.min.time())).replace(hour=self.end_of_day)

                blocks = sorted(busy.get(d, []))
                cur = day_s
                while cur + timedelta(hours=1) <= day_e:
                    slot_free = all(not (b_s < cur + timedelta(hours=1) and b_e > cur)
                                    for b_s, b_e, _ in blocks)
                    if slot_free:
                        free.setdefault(d, set()).add(cur)
                    cur += timedelta(hours=1)

            logging.info("Calendar OK in %.2fs  %s", time.time() - t0, url)
            return free

        except Exception as exc:
            logging.error("Calendar FAIL (%.2fs) %s – %s",
                          time.time() - t0, url, exc, exc_info=True)
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
        return "\n".join(formatted_output)+ "\n\n"

    def toggle_weekends(self):
        """Toggle weekend exclusion and update free slots"""
        self.exclude_weekends = not self.exclude_weekends
        self.update_free_slots()

    def paste_free_slots(self):
        """Copy free slots to clipboard and simulate paste."""
        logging.debug("Paste free slots triggered")
        try:
            if self.cached_free_slots:
                formatted_text = self.format_free_slots(self.cached_free_slots)

                # Store original clipboard content
                self.original_clipboard = pyperclip.paste()

                # Copy our text
                pyperclip.copy(formatted_text)
                logging.debug("Text copied to clipboard")

                # Use a longer delay before paste for macOS
                time.sleep(0.5 if platform.system() == 'Darwin' else 0.1)

                # Simulate paste - use a wrapper for pyautogui to improve reliability on macOS
                if platform.system() == 'Darwin':  # macOS
                    # Try paste multiple times with slight pauses to improve reliability
                    for attempt in range(3):
                        try:
                            pyautogui.hotkey('command', 'v')
                            logging.debug(f"Paste command sent (attempt {attempt + 1})")
                            # If we got here without error, break the loop
                            break
                        except Exception as e:
                            logging.error(f"Paste attempt {attempt + 1} failed: {e}")
                            if attempt < 2:  # Don't sleep after last attempt
                                time.sleep(0.2)
                else:  # Windows/Linux
                    pyautogui.hotkey('ctrl', 'v')
                    logging.debug("Paste command sent")

                # Restore original clipboard after a delay
                delay = 0.5
                if platform.system() == 'Darwin':
                    delay = 2.0  # Use a longer delay for macOS
                threading.Timer(delay, self.restore_clipboard).start()

                logging.debug("Paste operation completed")
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
                logging.debug("Calendar update completed successfully")
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
            logging.debug("Custom icon loaded successfully")
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
        logging.info("FreeTime application shutting down...")
        try:
            # Set shutting down flag
            self.is_shutting_down = True

            # Clean up keyboard hooks on Windows
            if platform.system() == "Windows":
                try:
                    import keyboard
                    keyboard.unhook_all()

                    # Additional Windows-specific cleanup
                    try:
                        # Force garbage collection
                        import gc
                        gc.collect()

                        # Release COM objects if any were created
                        import win32api
                        import win32con
                        # Post WM_QUIT to any hidden windows that might be lingering
                        win32api.PostQuitMessage(0)
                    except ImportError:
                        logging.debug("Windows-specific modules not available")
                except Exception as e:
                    logging.error(f"Error cleaning up Windows keyboard hooks: {e}")

            # Stop pynput listener on macOS
            elif platform.system() == "Darwin" and hasattr(self, '_keyboard_listener'):
                try:
                    self._keyboard_listener.stop()
                    self._keyboard_listener = None
                except Exception as e:
                    logging.error(f"Error stopping keyboard listener: {e}")

            # Restore clipboard if needed
            if hasattr(self, 'original_clipboard') and self.original_clipboard is not None:
                try:
                    pyperclip.copy(self.original_clipboard)
                    self.original_clipboard = None
                except Exception as e:
                    logging.error(f"Error restoring clipboard: {e}")

            # Stop icon if it exists
            if hasattr(self, 'icon') and self.icon:
                try:
                    logging.info("Stopping system tray icon...")
                    self.icon.stop()
                    self.icon = None
                except Exception as e:
                    logging.error(f"Error stopping icon: {e}")

            # Close any open windows
            try:
                for widget in self.root.winfo_children():
                    if isinstance(widget, tk.Toplevel) and widget.winfo_exists():
                        widget.destroy()
            except Exception as e:
                logging.error(f"Error closing windows: {e}")

        except Exception as e:
            logging.error(f"Error during cleanup: {e}")

        finally:
            # Always attempt to quit the root window
            try:
                logging.info("Quitting tkinter application...")
                self.root.quit()
                self.root.destroy()  # Explicitly destroy the root window
            except Exception as e:
                logging.error(f"Error quitting root window: {e}")

            logging.info("Application cleanup complete")

            # For Windows compiled app, ensure exit
            if platform.system() == "Windows" and getattr(sys, 'frozen', False):
                try:
                    # Force process termination as last resort
                    logging.info("Forcing process termination...")
                    import os
                    os._exit(0)  # Force immediate exit without cleanup
                except Exception:
                    pass

    def _ensure_cleanup(self):
        """Ensure cleanup happens no matter how the app exits"""
        if not hasattr(self, 'is_shutting_down') or not self.is_shutting_down:
            logging.info("Emergency cleanup triggered")
            self.cleanup()

        # Force exit any monitoring
        if platform.system() == "Darwin":
            try:
                from pynput import keyboard
                for t in threading.enumerate():
                    if isinstance(t, keyboard.Listener):
                        try:
                            t.stop()
                            logging.info(f"Stopped lingering keyboard thread: {t.name}")
                        except:
                            pass
            except:
                pass

    def show_about(self):
        """Show the About window"""
        try:
            if self.about_window is None or not tk.Toplevel.winfo_exists(self.about_window.window):
                #self.root.deiconify()  # Ensure root window exists
                self.about_window = AboutWindow(self.root, self.icon_image)
                #self.root.withdraw()  # Hide root window again
            else:
                self.about_window.window.lift()
                self.about_window.window.focus_force()
        except Exception as e:
            logging.error(f"Error showing About window: {e}")

    def run(self):
        """Run the application, with improved cleanup handling"""
        try:
            # Set initialize state
            self.is_shutting_down = False

            # Create and start update thread
            update_thread = threading.Thread(target=self.update_loop, daemon=True)
            update_thread.start()

            # Setup the hotkey monitoring
            self.setup_hotkey()

            # -------- pystray + tkinter MAIN LOOP HANDLING ---------
            system = platform.system()
            if system == "Darwin":
                # Register exit handler for macOS
                import atexit
                atexit.register(self._ensure_cleanup)

                # On macOS, pystray must run on the main thread, and tkinter also; avoid background threads for UI.
                self.icon = pystray.Icon(
                    "Calendar",
                    self.create_icon(),
                    "FreeTime",
                    menu=pystray.Menu(
                        pystray.MenuItem("About", lambda: self.show_about()),
                        pystray.MenuItem("Settings", lambda: self.show_settings()),
                        pystray.MenuItem("Update Now", lambda: self.update_free_slots()),
                        pystray.MenuItem("Exit", lambda: self.cleanup())
                    )
                )
                # Start pystray in the main thread
                self.icon.run_detached()  # non-blocking

                # Check if calendar URLs exist, if not show settings after a short delay
                if not self.calendar_urls:
                    self.root.after(1000, self.show_settings)

                # Now run tkinter mainloop (main thread)
                self.root.mainloop()

            else:
                # On Windows/Linux: Can run pystray in a separate thread
                self.icon = pystray.Icon(
                    "Calendar",
                    self.create_icon(),
                    "FreeTime",
                    menu=pystray.Menu(
                        pystray.MenuItem("About", lambda: self.root.after(0, self.show_about)),  # Safe for thread
                        pystray.MenuItem("Settings", lambda: self.root.after(0, self.show_settings)),
                        pystray.MenuItem("Update Now", lambda: self.update_free_slots()),
                        pystray.MenuItem("Exit", lambda: self.cleanup())
                    )
                )
                logging.debug("System tray icon initialized")

                if not self.calendar_urls:
                    self.root.after(1000, self.show_settings)

                # Start pystray in separate thread
                icon_thread = threading.Thread(target=self.icon.run, daemon=True)
                icon_thread.start()

                # Start tkinter mainloop
                self.root.mainloop()

        except Exception as e:
            logging.error(f"Fatal error in main loop: {e}")
            self._ensure_cleanup()  # Make sure cleanup happens even if we crash
            raise


if __name__ == "__main__":
    try:
        logging.info("=" * 50)
        logging.info("FreeTime Application Session Start")
        logging.info(f"Python version: {sys.version}")
        logging.info(f"Operating System: {platform.system()} {platform.version()}")

        app = CalendarApp()

        # Set up better process exit handling for Windows
        if platform.system() == "Windows":
            import atexit


            def ensure_exit():
                try:
                    if hasattr(app, 'cleanup'):
                        app.cleanup()
                    # Force exit for compiled Windows app
                    if getattr(sys, 'frozen', False):
                        import os
                        os._exit(0)
                except Exception:
                    pass


            atexit.register(ensure_exit)

        app.run()
    except KeyboardInterrupt:
        logging.info("Application terminated by user (KeyboardInterrupt)")
    except Exception as e:
        logging.error(f"Application error: {e}", exc_info=True)
    finally:
        # Final cleanup
        if 'app' in locals() and hasattr(app, 'cleanup'):
            app.cleanup()

        # Final forced exit for Windows compiled app
        if platform.system() == "Windows" and getattr(sys, 'frozen', False):
            try:
                import os

                os._exit(0)
            except Exception:
                pass