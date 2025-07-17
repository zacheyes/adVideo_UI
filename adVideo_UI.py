import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
import os
import subprocess
import shutil
from datetime import datetime
import sys
import json
import tkinter.font as tkFont
import requests
import filecmp
import threading
import pandas as pd
import stat # Import the 'stat' module for setting file permissions
import zipfile # MODIFIED: Added for zip file extraction

# --- Configuration ---
# Define your GitHub repository details
GITHUB_USERNAME = "zacheyes" # Placeholder, replace with actual if needed
GITHUB_REPO_NAME = "adVideo_UI" # Placeholder, replace with actual if needed
# This base URL points to the root of the 'main' branch for raw content.
# Ensure your scripts are directly in the root of the 'main' branch in your GitHub repo.
GITHUB_RAW_BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPO_NAME}/main/"

# --- GUI Script specific constants ---
# Renaming this GUI to be specific for video tasks
GUI_SCRIPT_FILENAME = "adVideo_UI.py"
UPDATE_IN_PROGRESS_MARKER = "adVideo_ui_update_in_progress.tmp"

# --- MODIFIED: Changed Mac launcher to launcher.zip ---
SCRIPT_FILENAMES = {
    "Video Renamer Script": "adVideo_renamer.py",
    "Bynder Video Metadata Export Script": "adVideo_metadataPrep.py",
    "Windows Launcher": "launcher.bat",
    "Mac Launcher": "launcher.zip", # MODIFIED for macOS
}

# --- MODIFIED: Changed URL to point to launcher.zip ---
GITHUB_SCRIPT_URLS = {
    GUI_SCRIPT_FILENAME: GITHUB_RAW_BASE_URL + GUI_SCRIPT_FILENAME,
    "adVideo_renamer.py": GITHUB_RAW_BASE_URL + "adVideo_renamer.py",
    "adVideo_metadataPrep.py": GITHUB_RAW_BASE_URL + "adVideo_metadataPrep.py",
    "launcher.bat": GITHUB_RAW_BASE_URL + "launcher.bat",
    "launcher.zip": GITHUB_RAW_BASE_URL + "launcher.zip", # MODIFIED for macOS
}

# Use a specific config file for this video UI to avoid conflicts with the larger GUI
CONFIG_FILE = "advideo_config.json"

# --- General Helper Functions ---

def _append_to_log(log_widget, text, is_stderr=False):
    log_widget.configure(state='normal')
    if is_stderr:
        log_widget.insert(tk.END, text, 'error')
    else:
        log_widget.insert(tk.END, text)
    log_widget.see(tk.END)
    log_widget.configure(state='disabled')

# --- Progress Bar Specific Helper Functions ---

def _prepare_progress_ui(progress_bar, progress_label, run_button_wrapper, progress_wrapper, initial_text):
    run_button_wrapper.grid_remove()
    progress_wrapper.grid(row=0, column=1)

    progress_bar.config(value=0, maximum=100)
    progress_bar.start() # Start indeterminate mode
    progress_label.config(text=initial_text)


def _update_progress_ui(progress_bar, progress_label, value, total_items=None):
    if progress_bar.cget('mode') != "determinate":
        progress_bar.config(mode="determinate") # Switch to determinate if not already

    if total_items is not None and total_items > 0:
        percent = (value / total_items) * 100
        progress_bar['value'] = percent
        progress_label.config(text=f"{percent:.1f}% ({value}/{total_items})")
    else:
        progress_bar['value'] = value
        progress_label.config(text=f"{value:.1f}%")

    progress_bar.update_idletasks()

def _on_process_complete_with_progress_ui(success, full_output, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, log_output_widget):
    if progress_bar:
        progress_bar.stop()
        progress_bar.config(mode="determinate", value=0)
    if progress_label:
        progress_label.config(text="")
    
    if progress_wrapper:
        progress_wrapper.grid_remove()
    if run_button_wrapper:
        run_button_wrapper.grid(row=0, column=1)

    if progress_bar and progress_bar.winfo_toplevel():
        progress_bar.winfo_toplevel().config(cursor="")
    elif log_output_widget and log_output_widget.winfo_toplevel():
        log_output_widget.winfo_toplevel().config(cursor="")

    if success:
        log_output_widget.insert(tk.END, "\nScript completed successfully.\n", 'success')
    else:
        log_output_widget.insert(tk.END, "\nScript failed. Please check the log above for errors.\n", 'error')
    log_output_widget.see(tk.END)
    
    if success:
        if success_callback:
            success_callback(full_output)
    else:
        if error_callback:
            error_callback(full_output)


# --- Run Script functions based on progress display needs ---

def _run_script_with_progress(script_full_path, args, log_output_widget, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, initial_progress_text):
    print("DEBUG (UI): Running script with progress bar.", file=sys.stderr)
    
    python_executable = sys.executable
    command = [python_executable, script_full_path]
    if args:
        command.extend(args)
    full_command_str = ' '.join(command)
    _append_to_log(log_output_widget, f"Executing subprocess command: {full_command_str}\n")

    log_output_widget.winfo_toplevel().after(0, lambda: _prepare_progress_ui(progress_bar, progress_label, run_button_wrapper, progress_wrapper, initial_progress_text))

    def _read_output_thread():
        process = None
        stdout_buffer = []
        stderr_buffer = []
        try:
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'

            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                         bufsize=1, universal_newlines=True, env=env)

            def read_stream(stream, buffer, is_stderr=False):
                for line in stream:
                    buffer.append(line)
                    log_output_widget.after(0, lambda log=log_output_widget, l=line: _append_to_log(log, l, is_stderr))
                    if line.startswith("PROGRESS:"):
                        try:
                            parts = line.split("PROGRESS:")[1].strip().split('/')
                            if len(parts) == 2:
                                value = float(parts[0])
                                total = float(parts[1])
                                progress_bar.after(0, lambda pb=progress_bar, pl=progress_label, val=value, tot=total: _update_progress_ui(pb, pl, val, tot))
                            else:
                                percent_val = float(parts[0])
                                progress_bar.after(0, lambda pb=progress_bar, pl=progress_label, val=percent_val: _update_progress_ui(pb, pl, val, 100))
                        except ValueError:
                            print(f"DEBUG (UI): Could not parse progress: {line.strip()}", file=sys.stderr)
                stream.close()

            stdout_thread = threading.Thread(target=read_stream, args=(process.stdout, stdout_buffer, False))
            stderr_thread = threading.Thread(target=read_stream, args=(process.stderr, stderr_buffer, True))
            
            stdout_thread.start()
            stderr_thread.start()

            stdout_thread.join()
            stderr_thread.join()

            process.wait()
            success = (process.returncode == 0)
            full_output = "".join(stdout_buffer) + "".join(stderr_buffer)
            log_output_widget.after(0, lambda: _on_process_complete_with_progress_ui(success, full_output, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, log_output_widget))

        except FileNotFoundError:
            error_msg = f"  Error: Python interpreter (or script) not found. Check paths and ensure Python is correctly installed and accessible.\n"
            log_output_widget.after(0, lambda: _append_to_log(log_output_widget, error_msg, is_stderr=True))
            log_output_widget.after(0, lambda: _on_process_complete_with_progress_ui(False, error_msg, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, log_output_widget))
        except Exception as e:
            error_msg = f"  An unexpected error occurred during subprocess execution: {e}\n"
            log_output_widget.after(0, lambda: _append_to_log(log_output_widget, error_msg, is_stderr=True))
            log_output_widget.after(0, lambda: _on_process_complete_with_progress_ui(False, error_msg, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, log_output_widget))

    subprocess_thread = threading.Thread(target=_read_output_thread)
    subprocess_thread.daemon = True
    subprocess_thread.start()
    return True, "Process started in background."


def _run_script_no_progress(script_full_path, args, log_output_widget, success_callback=None, error_callback=None):
    print("DEBUG (UI): Running script without progress bar.", file=sys.stderr)

    python_executable = sys.executable
    command = [python_executable, script_full_path]
    if args:
        command.extend(args)
    full_command_str = ' '.join(command)
    _append_to_log(log_output_widget, f"Executing subprocess command: {full_command_str}\n")

    try:
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        result = subprocess.run(command, capture_output=True, check=False, universal_newlines=True, env=env)
        
        stdout_str = result.stdout
        stderr_str = result.stderr
        
        full_output = stdout_str + stderr_str
        
        if stdout_str:
            _append_to_log(log_output_widget, "\n--- Script Output ---\n" + stdout_str)
        if stderr_str:
            _append_to_log(log_output_widget, "\n--- Script Errors (stderr) ---\n" + stderr_str, is_stderr=True)
        _append_to_log(log_output_widget, f"\nScript exited with return code: {result.returncode}\n")

        log_output_widget.winfo_toplevel().config(cursor="")
        if result.returncode == 0:
            if success_callback: success_callback(full_output)
            return True, full_output
        else:
            if error_callback: error_callback(full_output)
            return False, full_output

    except FileNotFoundError:
        error_msg = f"  Error: Python interpreter (or script) not found. Check paths and ensure Python is correctly installed and accessible.\n"
        _append_to_log(log_output_widget, error_msg, is_stderr=True)
        log_output_widget.winfo_toplevel().config(cursor="")
        if error_callback: error_callback("")
        return False, error_msg
    except Exception as e:
        error_msg = f"  An unexpected error occurred during subprocess execution: {e}\n"
        _append_to_log(log_output_widget, error_msg, is_stderr=True)
        log_output_widget.winfo_toplevel().config(cursor="")
        if error_callback: error_callback("")
        return False, error_msg

# --- Main Dispatcher Function for running scripts ---
def run_script_wrapper(script_full_path, is_python_script, args=None, log_output_widget=None,
                       progress_bar=None, progress_label=None, run_button_wrapper=None,
                       progress_wrapper=None, success_callback=None, error_callback=None,
                       initial_progress_text="Starting..."):
    
    print("DEBUG (UI): Entered run_script_wrapper function.", file=sys.stderr)

    if not os.path.exists(script_full_path):
        error_msg = f"Error: File not found at {script_full_path}\n"
        _append_to_log(log_output_widget, error_msg, is_stderr=True)
        log_output_widget.winfo_toplevel().config(cursor="")
        if error_callback: error_callback("")
        return False, error_msg

    if is_python_script:
        if progress_bar is not None and progress_label is not None and \
           run_button_wrapper is not None and progress_wrapper is not None:
            return _run_script_with_progress(script_full_path, args, log_output_widget,
                                              progress_bar, progress_label, run_button_wrapper,
                                              progress_wrapper, success_callback, error_callback,
                                              initial_progress_text)
        else:
            return _run_script_no_progress(script_full_path, args, log_output_widget,
                                           success_callback, error_callback)
    else:
        _append_to_log(log_output_widget, f"Opening file: {script_full_path}\n")
        try:
            if sys.platform == "win32":
                os.startfile(script_full_path)
            elif sys.platform == "darwin": # macOS
                # --- Ensure script is executable before opening on Mac ---
                if script_full_path.endswith('.sh'):
                    st = os.stat(script_full_path)
                    os.chmod(script_full_path, st.st_mode | stat.S_IEXEC)
                subprocess.Popen(["open", script_full_path])
            else: # Linux and other Unix-like
                if script_full_path.endswith('.sh'):
                    st = os.stat(script_full_path)
                    os.chmod(script_full_path, st.st_mode | stat.S_IEXEC)
                subprocess.Popen(["xdg-open", script_full_path])
            _append_to_log(log_output_widget, f"  File opened.\n")
            return True, f"Opened file: {script_full_path}"
        except Exception as e:
            _append_to_log(log_output_widget, f"  Error opening file: {e}\n", is_stderr=True)
            return False, f"Error opening file: {e}"


class Tooltip:
    def __init__(self, widget, text, bg_color, text_color):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.id = None
        self.x = 0
        self.y = 0
        self.bg_color = bg_color  
        self.text_color = text_color
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        self.x = self.widget.winfo_rootx() + 20  
        self.y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5  
        self.id = self.widget.after(500, self._display_tooltip)  

    def _display_tooltip(self):
        if self.tooltip_window:
            return
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)  
        self.tooltip_window.wm_geometry(f"+{self.x}+{self.y}")

        label = ttk.Label(self.tooltip_window, text=self.text, background=self.bg_color, relief=tk.SOLID, borderwidth=1,
                                          font=("Arial", 11), foreground=self.text_color, wraplength=400)
        label.pack(padx=5, pady=5)

    def hide_tooltip(self, event=None):
        if self.id:
            self.widget.after_cancel(self.id)
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

class AdVideoApp:
    def __init__(self, master):
        self.master = master
        master.title("Raymour & Flanigan Ad Video Tool")
        master.geometry("700x650")
        master.resizable(True, True)  

        self.current_theme = tk.StringVar(value="Light")
        self.style = ttk.Style()  
        
        self.base_font = tkFont.Font(family="Arial", size=10)
        self.header_font = tkFont.Font(family="Arial", size=12, weight="bold")
        self.log_font = tkFont.Font(family="Consolas", size=9)

        self._restarting_for_update = False

        self._initialize_logger_widget()

        self._apply_theme(self.current_theme.get())  

        self.scripts_root_folder = tk.StringVar(value=os.path.dirname(os.path.abspath(__file__)))
        self.last_update_timestamp = tk.StringVar(value="Last update: Never")
        self.gui_last_update_timestamp = tk.StringVar(value="Last GUI update: Never")
        
        self.video_renamer_spreadsheet_path = tk.StringVar(value="")
        self.video_renamer_folder_path = tk.StringVar(value="")

        self.bynder_video_metadata_spreadsheet_path = tk.StringVar(value="")
        self.bynder_video_metadata_folder_path = tk.StringVar(value="")

        self.link_to_wrike_project_ui = tk.StringVar(value="")
        current_year = datetime.now().year
        self.available_years = [""] + [str(y) for y in range(current_year - 2, current_year + 3)]
        self.year_ui = tk.StringVar(value="")
        
        self.sub_initiative_ui = tk.StringVar(value="")
        self.location_type_options = ["", "Clearance Center", "Corporate Location", "Customer Service Center (CSC)", "Distribution Center", "Full Line", "Hybrid", "Outlet"]
        self.location_type_ui = tk.StringVar(value="")


        self.log_expanded = False

        self._create_widgets()
        self._load_configuration()

        self.log_print(f"UI launched with Python {sys.version.split(' ')[0]} from: {sys.executable}\n")
        self.log_print(f"Operating System: {sys.platform}\n")
        self.log_print("Ad Video Tool initialized. Please select paths and run operations.\n")

        self.master.after(100, self._handle_startup_update_check)

        master.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _initialize_logger_widget(self):
        self.log_text_early_placeholder = scrolledtext.ScrolledText(self.master, width=1, height=1, state='disabled')
        
        def custom_print(*args, **kwargs):
            text = " ".join(map(str, args)) + kwargs.get('end', '\n')
            if hasattr(self, 'log_text') and self.log_text.winfo_exists():
                self.log_text.configure(state='normal')
                self.log_text.insert(tk.END, text)
                self.log_text.see(tk.END)
                self.log_text.configure(state='disabled')
            else:
                print(text, end='')  

        self.log_print = custom_print
        self.log_text = self.log_text_early_placeholder


    def _on_closing(self):
        if not self._restarting_for_update:
            self._save_configuration()
        self.master.destroy()

    def _save_configuration(self):
        config_data = {
            "theme": self.current_theme.get(),
            "scripts_root_folder": self.scripts_root_folder.get(),
            "last_update": self.last_update_timestamp.get(),
            "gui_last_update": self.gui_last_update_timestamp.get(),
        }
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config_data, f, indent=4)
            self.log_print("Configuration saved successfully.\n")
        except Exception as e:
            self.log_print(f"Error saving configuration: {e}\n")

    def _load_configuration(self):
        """Loads specified configuration items from the JSON file."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    config_data = json.load(f)
                
                self.scripts_root_folder.set(config_data.get("scripts_root_folder", os.path.dirname(os.path.abspath(__file__))))
                loaded_theme = config_data.get("theme", "Light")
                self.current_theme.set(loaded_theme)
                self._apply_theme(loaded_theme)

                last_update_from_config = config_data.get("last_update", "Last update: Never")
                if not last_update_from_config.startswith("Last update:"):
                    self.last_update_timestamp.set(f"Last update: {last_update_from_config}")
                else:
                    self.last_update_timestamp.set(last_update_from_config)

                gui_last_update_from_config = config_data.get("gui_last_update", "Last GUI update: Never")
                if not gui_last_update_from_config.startswith("Last GUI update:"):
                    self.gui_last_update_timestamp.set(f"Last GUI update: {gui_last_update_from_config}")
                else:
                    self.gui_last_update_timestamp.set(gui_last_update_from_config)

                self.log_print("Configuration loaded successfully.\n")
            except json.JSONDecodeError as e:
                self.log_print(f"Error reading configuration file (JSON format issue): {e}\n")
            except Exception as e:
                self.log_print(f"Error loading configuration: {e}\n")
        else:
            self.log_print("No existing configuration file found. Using default paths.\n")

        self._setup_initial_state()


    def _apply_theme(self, theme_name):
        self.current_theme.set(theme_name)

        self.RF_PURPLE_BASE = "#4f245e"  
        self.RF_WHITE_BASE = "#FFFFFF"  

        if theme_name == "Dark":
            self.primary_bg = "#2B2B2B"  
            self.secondary_bg = "#3C3C3C"  
            self.text_color = "#E0E0E0"  
            self.header_text_color = "#FFFFFF"  
            self.accent_color = self.RF_PURPLE_BASE
            self.border_color = "#555555"  
            self.log_bg = "#1E1E1E"  
            self.log_text_color = "#CCCCCC"  
            self.trough_color = "#555555"  
            self.slider_color = "#888888"  
            self.checkbox_indicator_off = "#3C3C3C"  
            self.checkbox_indicator_on = self.accent_color
            self.checkbox_hover_bg = "#505050"  
            self.radiobutton_hover_bg = "#505050"

        else: # Light Theme
            self.primary_bg = "#F0F0F0"  
            self.secondary_bg = "#FFFFFF"  
            self.text_color = "#333333"  
            self.header_text_color = self.RF_PURPLE_BASE  
            self.accent_color = self.RF_PURPLE_BASE
            self.border_color = "#CCCCCC"  
            self.log_bg = "#E8E8E8"  
            self.log_text_color = "#444444"  
            self.trough_color = "#E0E0E0"  
            self.slider_color = "#BBBBBB"  
            self.checkbox_indicator_off = "#E0E0E0"
            self.checkbox_indicator_on = self.accent_color
            self.checkbox_hover_bg = "#E0E0E0"
            self.radiobutton_hover_bg = "#E0E0E0"
            
        self.master.config(bg=self.primary_bg)
        if hasattr(self, 'canvas'):  
            self.canvas.config(bg=self.primary_bg)

        self.style.theme_use("clam")  

        self.style.configure('.',
                             font=self.base_font,
                             background=self.primary_bg,
                             foreground=self.text_color)
        
        self.style.configure('TFrame',
                             background=self.primary_bg)
        
        self.style.configure('SectionFrame.TFrame',
                             background=self.secondary_bg,
                             borderwidth=1,
                             relief="solid",  
                             padding=0)  

        self.style.configure('TLabel',
                             background=self.primary_bg,
                             foreground=self.text_color)
        
        self.style.configure('Header.TLabel',
                             font=self.header_font,
                             foreground=self.header_text_color,
                             background=self.secondary_bg)  

        self.style.configure('TButton',
                             background=self.accent_color,
                             foreground=self.RF_WHITE_BASE,  
                             font=self.base_font,
                             relief='flat',
                             padding=5)
        self.style.map('TButton',
                       background=[('active', self._shade_color(self.accent_color, -0.1))],  
                       foreground=[('active', self.RF_WHITE_BASE)])  

        self.style.configure('TEntry',
                             fieldbackground=self.secondary_bg,
                             foreground=self.text_color,
                             borderwidth=1,
                             relief="solid")
        
        self.style.configure('TScrollbar',
                             troughcolor=self.trough_color,
                             background=self.slider_color,
                             bordercolor=self.trough_color,
                             arrowcolor=self.text_color)
        self.style.map('TScrollbar',
                       background=[('active', self._shade_color(self.slider_color, -0.1))])

        self.style.configure('TNotebook',
                             background=self.primary_bg,
                             borderwidth=0)
        self.style.configure('TNotebook.Tab',
                             background=self._shade_color(self.primary_bg, -0.05),  
                             foreground=self.text_color,
                             font=self.base_font,
                             padding=[5, 2])
        self.style.map('TNotebook.Tab',
                       background=[('selected', self.accent_color)],
                       foreground=[('selected', self.RF_WHITE_BASE)],  
                       expand=[('selected', [1, 1, 1, 0])])  

        self.style.configure('TRadiobutton',
                             background=self.primary_bg,
                             foreground=self.text_color,
                             font=self.base_font,
                             indicatorcolor=self.accent_color)
        self.style.map('TRadiobutton',
                       background=[('active', self.radiobutton_hover_bg)],
                       foreground=[('active', self.text_color)],
                       indicatorcolor=[('selected', self.accent_color), ('!selected', self.checkbox_indicator_off)])
        
        self.style.configure('TCheckbutton',
                             background=self.primary_bg,
                             foreground=self.text_color,
                             font=self.base_font,
                             indicatorcolor=self.checkbox_indicator_off)
        self.style.map('TCheckbutton',
                       background=[('active', self.checkbox_hover_bg)],
                       foreground=[('active', self.text_color)],
                       indicatorcolor=[('selected', self.checkbox_indicator_on), ('!selected', self.checkbox_indicator_off)])

        self.style.configure('TSeparator', background=self.border_color, relief='solid', sashrelief='solid', sashwidth=3)
        self.style.layout('TSeparator',
                                [('TSeparator.separator', {'sticky': 'nswe'})])

        self.style.configure('TCombobox',
                             fieldbackground=self.secondary_bg,  
                             background=self.primary_bg,  
                             foreground=self.text_color,
                             arrowcolor=self.text_color)
        self.style.map('TCombobox',
                       fieldbackground=[('readonly', self.secondary_bg)],
                       background=[('readonly', self.primary_bg)],
                       foreground=[('readonly', self.text_color)],
                       selectbackground=[('readonly', self._shade_color(self.secondary_bg, -0.05))],  
                       selectforeground=[('readonly', self.text_color)])  

        if hasattr(self, 'log_text'):
            self.log_text.config(bg=self.log_bg, fg=self.log_text_color,
                                 insertbackground=self.log_text_color,
                                 selectbackground=self.accent_color,
                                 selectforeground=self.RF_WHITE_BASE)
            self.log_text.tag_config('error', foreground='#FF6B6B')
            self.log_text.tag_config('success', foreground='#6BFF6B')
        
        self._update_all_widget_colors()  

    def _shade_color(self, hex_color, percent):
        """Shades a hex color by a given percentage. Positive percent for lighter, negative for darker."""
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        new_rgb = []
        for color_val in rgb:
            new_val = color_val * (1 + percent)
            new_val = max(0, min(255, int(new_val)))
            new_rgb.append(new_val)
            
        return '#%02x%02x%02x' % tuple(new_rgb)

    def _update_all_widget_colors(self):
        for widget in self.master.winfo_children():
            self._update_widget_color_recursive(widget)

    def _update_widget_color_recursive(self, widget):
        try:
            if hasattr(widget, 'config'):
                if 'background' in widget.config():
                    if isinstance(widget, ttk.Label) and widget.cget('style') == 'Header.TLabel':
                        widget.config(background=self.secondary_bg)  
                    else:
                        widget.config(background=self.primary_bg)
                if 'foreground' in widget.config():
                    if isinstance(widget, ttk.Label) and widget.cget('style') == 'Header.TLabel':
                        widget.config(foreground=self.header_text_color)
                    else:
                        widget.config(foreground=self.text_color)
            
            if isinstance(widget, tk.Canvas):
                widget.config(bg=self.primary_bg)
            elif isinstance(widget, scrolledtext.ScrolledText):
                widget.config(bg=self.log_bg, fg=self.log_text_color,
                              insertbackground=self.log_text_color,
                              selectbackground=self.accent_color,
                              selectforeground=self.RF_WHITE_BASE)

        except tk.TclError:
            pass  

        for child_widget in widget.winfo_children():
            self._update_widget_color_recursive(child_widget)
            
    def _on_theme_change(self, event=None):
        selected_theme = self.current_theme.get()
        self._apply_theme(selected_theme)
        self._save_configuration()  

    def _toggle_log_size(self):
        if self.log_expanded:  
            self.log_text.pack_forget()  
            self.toggle_log_button.config(text="▲")  
            self.master.grid_rowconfigure(2, weight=0)  
            self.log_wrapper_frame.config(height=50)  
            self.log_expanded = False  
        else:  
            self.log_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)  
            self.toggle_log_button.config(text="▼")  
            self.master.grid_rowconfigure(2, weight=1)  
            self.log_expanded = True  
        self._save_configuration()  

    def _browse_scripts_root_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.scripts_root_folder.set(folder_path)
            self._save_configuration()

    def _browse_folder(self, string_var):
        folder_path = filedialog.askdirectory()
        if folder_path:
            string_var.set(folder_path)

    def _browse_file(self, string_var, file_type):
        if file_type == "spreadsheet":
            file_types = [("Spreadsheet files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        else:
            file_types = [("All files", "*.*")]

        file_path = filedialog.askopenfilename(filetypes=file_types)
        if file_path:
            string_var.set(file_path)

    def _ensure_dir(self, path):
        """Ensures the directory for a given path exists."""
        directory = os.path.dirname(path) if os.path.basename(path) and '.' in os.path.basename(path) else path
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
            self.log_print(f"  Created directory: {directory}")

    # MODIFIED: New helper method for extracting zip and setting permissions
    def _extract_and_permission_launcher(self, zip_path, extract_folder):
        """Extracts the launcher.zip and sets permissions on launcher.command."""
        self.log_print(f"  Processing '{os.path.basename(zip_path)}'...")
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                self.log_print(f"  Extracting all contents from '{os.path.basename(zip_path)}'...")
                zip_ref.extractall(extract_folder)
                self.log_print(f"  Successfully extracted to '{extract_folder}'.")

            # Set execute permissions on the extracted launcher.command
            extracted_sh_path = os.path.join(extract_folder, "launcher.command")
            if os.path.exists(extracted_sh_path):
                st = os.stat(extracted_sh_path)
                # Sets permissions to rwxr-xr-x (read/write/execute for owner, read/execute for group/others)
                os.chmod(extracted_sh_path, st.st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)
                self.log_print(f"  Set execute permissions for 'launcher.command'.\n")
            else:
                self.log_print(f"  WARNING: 'launcher.command' not found after extraction. Check the contents of the zip file.\n", is_stderr=True)
        
        except zipfile.BadZipFile:
            self.log_print(f"  ERROR: '{os.path.basename(zip_path)}' is not a valid zip file.\n", is_stderr=True)
        except Exception as e:
            self.log_print(f"  ERROR processing launcher zip: {e}\n", is_stderr=True)


    # MODIFIED: Rewritten to handle zip extraction
    def _download_and_compare_file(self, display_name, filename, download_url, local_target_folder):
        local_full_path = os.path.join(local_target_folder, filename)
        temp_file_path = local_full_path + ".tmp"
        
        self.log_print(f"Checking {display_name} ({filename})....")
        self.log_print(f"  Local path: {local_full_path}")
        self.log_print(f"  Download URL: {download_url}")

        try:
            response = requests.get(download_url, stream=True)
            response.raise_for_status()

            self._ensure_dir(local_full_path)

            with open(temp_file_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            status = ""
            # Check if the downloaded file is identical to the local one
            if os.path.exists(local_full_path) and filecmp.cmp(local_full_path, temp_file_path, shallow=False):
                self.log_print(f"  '{filename}' is already up to date. No action needed.")
                os.remove(temp_file_path)
                # EVEN IF SKIPPED: For Mac launcher, ensure it's extracted and executable.
                if filename == "launcher.zip" and sys.platform == "darwin":
                    self._extract_and_permission_launcher(local_full_path, local_target_folder)
                else:
                    self.log_print("\n")
                return "skipped"
            
            # If new or updated, replace the local file
            if os.path.exists(local_full_path):
                self.log_print(f"  New version of '{filename}' found. Updating...")
                status = "updated"
            else:
                self.log_print(f"  '{filename}' not found locally. Downloading new file...")
                status = "downloaded"

            shutil.move(temp_file_path, local_full_path)
            self.log_print(f"  '{filename}' {status} successfully!")

            # If it's the Mac launcher zip, extract it and set permissions.
            if filename == "launcher.zip" and sys.platform == "darwin":
                self._extract_and_permission_launcher(local_full_path, local_target_folder)
            else:
                # Keep consistent spacing for other files
                self.log_print("\n")

            return status

        except requests.exceptions.RequestException as e:
            self.log_print(f"  ERROR downloading '{filename}' from '{download_url}': {e}\n", is_stderr=True)
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
            return "error"
        except Exception as e:
            self.log_print(f"  An unexpected ERROR occurred while updating '{filename}': {e}\n", is_stderr=True)
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
            return "error"

    def _update_all_scripts(self):
            scripts_folder = self.scripts_root_folder.get()
            if not scripts_folder or not os.path.isdir(scripts_folder):
                messagebox.showerror("Error", "Please set a valid 'Local Scripts Folder' first.")
                self.log_print("Update aborted: 'Local Scripts Folder' is not set or invalid.\n", is_stderr=True)
                return

            self.log_print("\n--- Starting All Scripts & Files Update Process ---")
            self.log_print(f"Using root folder: {scripts_folder}\n")

            updated_count = 0
            downloaded_count = 0
            skipped_count = 0
            error_count = 0
            launcher_or_reqs_updated = False

            # --- MODIFIED: Platform-aware file checking for .zip on Mac ---
            files_to_check = {}
            for display_name, filename in SCRIPT_FILENAMES.items():
                is_launcher = "launcher" in filename.lower()
                # If it's not a launcher, always add it
                if not is_launcher:
                    files_to_check[display_name] = filename
                # If it's a launcher, check the platform
                else:
                    if sys.platform == "win32" and filename.endswith(".bat"):
                        files_to_check[display_name] = filename
                    elif sys.platform == "darwin" and filename.endswith(".zip"): # MODIFIED
                        files_to_check[display_name] = filename
            
            self.log_print(f"Platform '{sys.platform}' detected. Checking relevant files...\n")

            for display_name, filename in files_to_check.items():
                if filename not in GITHUB_SCRIPT_URLS:
                    continue
                
                github_url = GITHUB_SCRIPT_URLS[filename]
                status = self._download_and_compare_file(display_name, filename, github_url, scripts_folder)
                
                if status in ["updated", "downloaded"]:
                    if status == "updated":
                        updated_count += 1
                    else: # downloaded
                        downloaded_count += 1
                    
                    # --- MODIFIED: Check for launcher zip file ---
                    if filename in ["launcher.bat", "launcher.zip"]:
                        launcher_or_reqs_updated = True
                
                elif status == "skipped":
                    skipped_count += 1
                elif status == "error":
                    error_count += 1
            
            self.log_print("\n--- All Update Processes Complete ---")
            
            summary_message_parts = []
            if updated_count > 0:
                summary_message_parts.append(f"Updated {updated_count} file(s).")
            if downloaded_count > 0:
                summary_message_parts.append(f"Newly downloaded {downloaded_count} file(s).")
            
            final_message = "Update process finished."
            if summary_message_parts:
                final_message += "\n\n" + "\n".join(summary_message_parts)
            elif error_count == 0:
                final_message = "All local files are already up to date."

            if error_count > 0:
                final_message += f"\n\nEncountered {error_count} error(s). Please check the log."

            messagebox.showinfo("Update Complete", final_message)
            
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.last_update_timestamp.set(f"Last update: {current_time}")
            self._save_configuration()

    def _check_for_gui_update(self):
        """Checks for a new version of the GUI script and updates/restarts if available."""
        self.log_print("\n--- Checking for GUI script update ---")
        local_gui_path = os.path.abspath(sys.argv[0])
        github_url = GITHUB_SCRIPT_URLS.get(GUI_SCRIPT_FILENAME)
        
        if not github_url:
            self.log_print("Error: GUI script URL not found in configuration.", is_stderr=True)
            messagebox.showerror("Update Error", "GUI script URL not configured.")
            return

        temp_download_path = local_gui_path + ".new_version_tmp"

        try:
            self.log_print(f"Downloading latest GUI from: {github_url}")
            response = requests.get(github_url, stream=True)
            response.raise_for_status()

            downloaded_content = response.content
            with open(temp_download_path, 'wb') as f:
                f.write(downloaded_content)
            
            gui_updated = False
            if os.path.exists(local_gui_path):
                with open(local_gui_path, 'rb') as f:
                    local_content = f.read()
                
                if local_content == downloaded_content:
                    self.log_print("GUI script is already up to date. No update needed.\n")
                    messagebox.showinfo("Update Check", "The GUI is already up to date!")
                else:
                    self.log_print("New version of GUI script found. Applying update...")
                    shutil.copy(temp_download_path, local_gui_path)
                    gui_updated = True
            else:
                self.log_print("GUI script not found locally. Downloading new script...")
                shutil.copy(temp_download_path, local_gui_path)
                gui_updated = True

            if gui_updated:
                with open(UPDATE_IN_PROGRESS_MARKER, 'w') as f:
                    f.write(str(os.getpid()))

                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.gui_last_update_timestamp.set(f"Last GUI update: {current_time}")
                self._save_configuration()

                self.log_print("GUI script updated successfully. Restarting application...\n")
                messagebox.showinfo("Update Complete", "The GUI has been updated. The application will now restart to apply changes.")
                
                try:
                    if os.path.exists(temp_download_path):
                        os.remove(temp_download_path)
                        self.log_print(f"Cleaned up temporary file: {temp_download_path}\n")
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary download file before restart: {e}", is_stderr=True)

                self._restarting_for_update = True
                python = sys.executable
                os.execl(python, python, *sys.argv)
            
        except requests.exceptions.RequestException as e:
            self.log_print(f"Error checking/downloading GUI update: {e}\n", is_stderr=True)
            messagebox.showerror("Update Error", f"Failed to check for GUI update: {e}")
        except Exception as e:
            self.log_print(f"An unexpected error occurred during GUI update: {e}\n", is_stderr=True)
            messagebox.showerror("Update Error", f"An unexpected error occurred: {e}")
        finally:
            if os.path.exists(temp_download_path):
                try:
                    os.remove(temp_download_path)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary download file: {e}", is_stderr=True)

    def _handle_startup_update_check(self):
        """Checks for and cleans up the update marker file on startup."""
        if os.path.exists(UPDATE_IN_PROGRESS_MARKER):
            try:
                os.remove(UPDATE_IN_PROGRESS_MARKER)
                self.log_print("GUI update completed successfully. Welcome back!\n", 'success')
            except Exception as e:
                self.log_print(f"Warning: Could not remove update marker file: {e}\n", is_stderr=True)

    def _run_video_renamer_script(self):
        spreadsheet_path = self.video_renamer_spreadsheet_path.get()
        folder_path = self.video_renamer_folder_path.get()
        script_name = SCRIPT_FILENAMES["Video Renamer Script"]
        script_full_path = os.path.join(self.scripts_root_folder.get(), script_name)

        if not os.path.exists(script_full_path):
            messagebox.showerror("Error", f"Video Renamer Script not found: {script_full_path}\n"
                                         f"Please ensure '{script_name}' is in your scripts folder.")
            return
        if not spreadsheet_path or not os.path.exists(spreadsheet_path) or not spreadsheet_path.lower().endswith(('.xlsx', '.xls', '.csv')):
            messagebox.showerror("Input Error", "Please select a valid Spreadsheet (.xlsx, .xls, or .csv).")
            return
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("Input Error", "Please select a valid Folder of Video Files.")
            return

        self.log_print(f"\n--- Running Video Renamer Script ({script_name}) ---")
        self.log_print(f"Spreadsheet: {spreadsheet_path}")
        self.log_print(f"Video Folder: {folder_path}")
        
        args = [
            "--spreadsheet", spreadsheet_path,
            "--video_folder", folder_path
        ]
            
        self.run_video_renamer_button.config(state='disabled')
        _prepare_progress_ui(self.video_renamer_progress_bar, self.video_renamer_progress_label,
                             self.video_renamer_run_button_wrapper, self.video_renamer_progress_wrapper,
                             initial_text="Renaming videos...")

        def success_callback(output):
            self.run_video_renamer_button.config(state='normal')
            messagebox.showinfo("Success", "Video Renamer script completed successfully!")
            self.log_print("Video Renamer completed.\n", 'success')

        def error_callback(output):
            self.run_video_renamer_button.config(state='normal')
            messagebox.showerror("Error", "Video Renamer script failed. Please check the log for details.")
            self.log_print("Video Renamer failed.\n", 'error')

        run_script_wrapper(script_full_path, True, args, self.log_text, 
                           self.video_renamer_progress_bar, self.video_renamer_progress_label,
                           self.video_renamer_run_button_wrapper, self.video_renamer_progress_wrapper,
                           success_callback, error_callback,
                           initial_progress_text="Renaming videos...")


    def _run_bynder_video_metadata_export(self):
        spreadsheet_path = self.bynder_video_metadata_spreadsheet_path.get()
        folder_path = self.bynder_video_metadata_folder_path.get()
        
        wrike_link = self.link_to_wrike_project_ui.get().strip()
        year = self.year_ui.get().strip()
        sub_initiative = self.sub_initiative_ui.get().strip()
        location_type = self.location_type_ui.get().strip()

        script_name = SCRIPT_FILENAMES["Bynder Video Metadata Export Script"]
        script_full_path = os.path.join(self.scripts_root_folder.get(), script_name)

        if not os.path.exists(script_full_path):
            messagebox.showerror("Error", f"Bynder Video Metadata Export Script not found: {script_full_path}\n"
                                         f"Please ensure '{script_name}' is in your scripts folder.")
            return
        if not spreadsheet_path or not os.path.exists(spreadsheet_path) or not spreadsheet_path.lower().endswith(('.xlsx', '.xls', '.csv')):
            messagebox.showerror("Input Error", "Please select a valid Spreadsheet (.xlsx, .xls, or .csv).")
            return
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("Input Error", "Please select a valid Folder of Renamed Assets (Videos).")
            return
        
        self.log_print(f"\n--- Running Bynder Video Metadata Export Script ({script_name}) ---")
        self.log_print(f"Spreadsheet: {spreadsheet_path}")
        self.log_print(f"Assets Folder: {folder_path}")
        self.log_print(f"UI provided values: Wrike='{wrike_link}', Year='{year}', Sub-Initiative='{sub_initiative}', Location Type='{location_type}'")

        args = [
            "--spreadsheet", spreadsheet_path,
            "--assets_folder", folder_path,
        ]
        
        if wrike_link:
            args.extend(["--wrike_link", wrike_link])
        if year:
            args.extend(["--year", year])
        if sub_initiative:
            args.extend(["--sub_initiative", sub_initiative])
        if location_type:
            args.extend(["--location_type", location_type])
            
        self.run_bynder_video_metadata_export_button.config(state='disabled')
        _prepare_progress_ui(self.bynder_video_metadata_export_progress_bar, self.bynder_video_metadata_export_progress_label,
                             self.bynder_video_metadata_export_run_button_wrapper, self.bynder_video_metadata_export_progress_wrapper,
                             initial_text="Exporting Bynder metadata...")

        def success_callback(output):
            self.run_bynder_video_metadata_export_button.config(state='normal')
            messagebox.showinfo("Success", "Bynder Video Metadata Export script completed successfully!\n"
                                         "The metadata CSV should be in your Downloads folder.")
            self.log_print("Bynder Video Metadata Export completed.\n", 'success')

        def error_callback(output):
            self.run_bynder_video_metadata_export_button.config(state='normal')
            messagebox.showerror("Error", "Bynder Video Metadata Export script failed. Please check the log for details.")
            self.log_print("Bynder Video Metadata Export failed.\n", 'error')

        run_script_wrapper(script_full_path, True, args, self.log_text, 
                           self.bynder_video_metadata_export_progress_bar, self.bynder_video_metadata_export_progress_label,
                           self.bynder_video_metadata_export_run_button_wrapper, self.bynder_video_metadata_export_progress_wrapper,
                           success_callback, error_callback,
                           initial_progress_text="Exporting Bynder metadata...")


    def _setup_initial_state(self):
        """Sets up the initial state of the GUI elements."""
        if not self.log_expanded:  
            self.log_text.pack_forget()  
            self.toggle_log_button.config(text="▲")  
            self.master.grid_rowconfigure(2, weight=0)  
            self.log_wrapper_frame.config(height=50)  
        
    def _create_widgets(self):
        self.master.grid_rowconfigure(0, weight=0)
        self.master.grid_rowconfigure(1, weight=2)
        self.master.grid_rowconfigure(2, weight=1)
        self.master.grid_rowconfigure(3, weight=0)
        self.master.grid_columnconfigure(0, weight=1)  

        top_bar_frame = ttk.Frame(self.master, style='TFrame')
        top_bar_frame.grid(row=0, column=0, padx=(10, 10), pady=(2, 2), sticky="new")  
        
        top_bar_frame.grid_columnconfigure(0, weight=1)
        top_bar_frame.grid_columnconfigure(1, weight=1)
        top_bar_frame.grid_columnconfigure(2, weight=0)

        update_all_scripts_section = ttk.Frame(top_bar_frame, style='TFrame')
        update_all_scripts_section.grid(row=0, column=0, padx=(0, 10), sticky="w")
        update_all_scripts_section.grid_columnconfigure(0, weight=1)
        
        self.update_all_scripts_button = ttk.Button(update_all_scripts_section, text="Update All Scripts", command=self._update_all_scripts, style='TButton')
        self.update_all_scripts_button.pack(fill="x", expand=True)
        Tooltip(self.update_all_scripts_button, "Checks GitHub for updated versions of Python scripts and downloads them to your local scripts folder if newer versions are available. If a script is missing, it will download it. Also checks for the ExifTool bundle.", self.secondary_bg, self.text_color)
        
        self.last_update_label = ttk.Label(update_all_scripts_section, textvariable=self.last_update_timestamp, style='TLabel')
        self.last_update_label.pack(pady=(2,0))


        update_gui_section = ttk.Frame(top_bar_frame, style='TFrame')
        update_gui_section.grid(row=0, column=1, padx=(0, 10), sticky="w")
        update_gui_section.grid_columnconfigure(0, weight=1)

        self.check_gui_update_button = ttk.Button(update_gui_section, text="Update GUI", command=self._check_for_gui_update, style='TButton')
        self.check_gui_update_button.pack(fill="x", expand=True)
        Tooltip(self.check_gui_update_button, "Checks for and applies updates to this GUI application itself, then restarts.", self.secondary_bg, self.text_color)
        
        self.gui_last_update_label = ttk.Label(update_gui_section, textvariable=self.gui_last_update_timestamp, style='TLabel')
        self.gui_last_update_label.pack(pady=(2,0))


        theme_frame = ttk.Frame(top_bar_frame, style='TFrame')
        theme_frame.grid(row=0, column=2, sticky="e")  
        
        self.theme_label = ttk.Label(theme_frame, text="Theme:", style='TLabel')
        self.theme_label.pack(side="left", padx=(0, 5))
        
        self.theme_selector = ttk.Combobox(theme_frame, textvariable=self.current_theme,  
                                           values=["Light", "Dark"], state="readonly", width=6)
        self.theme_selector.pack(side="left")
        self.theme_selector.bind("<<ComboboxSelected>>", self._on_theme_change)
        

        container = ttk.Frame(self.master)
        container.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")  

        self.canvas = tk.Canvas(container, highlightthickness=0, bg=self.primary_bg)  
        self.canvas.pack(side="left", fill="both", expand=True)  

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.scrollable_frame = ttk.Frame(self.canvas, style='TFrame')  
        canvas_frame_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        def on_frame_configure(event):
            canvas_width = event.width
            self.canvas.itemconfig(canvas_frame_id, width=canvas_width)  
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        self.scrollable_frame.bind("<Configure>", on_frame_configure)
        self.canvas.bind("<Configure>", lambda event: self.canvas.itemconfig(canvas_frame_id, width=event.width))

        def _on_mouse_wheel(event):
            if sys.platform == "darwin":  
                self.canvas.yview_scroll(int(-1*(event.delta)), "units")
            else:  
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        self.canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

        row_counter = 0  

        # SECTION: Local Scripts Folder Path
        scripts_folder_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        scripts_folder_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_scripts = ttk.Frame(scripts_folder_wrapper_frame, style='TFrame')
        header_sub_frame_scripts.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_scripts = ttk.Label(header_sub_frame_scripts, text="Local Scripts Folder", style='Header.TLabel')
        header_label_scripts.pack(side="left", padx=(0, 5))
        info_label_scripts = ttk.Label(header_sub_frame_scripts, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_scripts, "This is the local folder where all your Python scripts are located. The application will look for and save scripts in this directory.", self.secondary_bg, self.text_color)  
        info_label_scripts.pack(side="left", anchor="center")

        scripts_folder_frame = ttk.Frame(scripts_folder_wrapper_frame, style='TFrame')
        scripts_folder_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        scripts_folder_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(scripts_folder_frame, text="Path to Scripts Folder:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(scripts_folder_frame, textvariable=self.scripts_root_folder, width=40, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(scripts_folder_frame, text="Browse", command=self._browse_scripts_root_folder, style='TButton').grid(row=0, column=2, padx=5, pady=5)

        # SECTION: Video Renamer Tool
        video_renamer_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        video_renamer_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_video_renamer = ttk.Frame(video_renamer_wrapper_frame, style='TFrame')
        header_sub_frame_video_renamer.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_video_renamer = ttk.Label(header_sub_frame_video_renamer, text="Video Renamer Tool", style='Header.TLabel')
        header_label_video_renamer.pack(side="left", padx=(0, 5))
        info_label_video_renamer = ttk.Label(header_sub_frame_video_renamer, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_video_renamer, "Renames video files based on a spreadsheet. Matches files by 'Description' column and appends '-AD ID'.", self.secondary_bg, self.text_color)  
        info_label_video_renamer.pack(side="left", anchor="center")

        video_renamer_frame = ttk.Frame(video_renamer_wrapper_frame, style='TFrame')
        video_renamer_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        video_renamer_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(video_renamer_frame, text="Spreadsheet:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(video_renamer_frame, textvariable=self.video_renamer_spreadsheet_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(video_renamer_frame, text="Browse File", command=lambda: self._browse_file(self.video_renamer_spreadsheet_path, "spreadsheet"), style='TButton').grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(video_renamer_frame, text="Folder of Video Files:", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(video_renamer_frame, textvariable=self.video_renamer_folder_path, width=45, style='TEntry').grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(video_renamer_frame, text="Browse Folder", command=lambda: self._browse_folder(self.video_renamer_folder_path), style='TButton').grid(row=1, column=2, padx=5, pady=5)

        self.video_renamer_run_control_frame = ttk.Frame(video_renamer_frame, style='TFrame')
        self.video_renamer_run_control_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky="ew")
        self.video_renamer_run_control_frame.grid_columnconfigure(0, weight=1)
        self.video_renamer_run_control_frame.grid_columnconfigure(1, weight=0)
        self.video_renamer_run_control_frame.grid_columnconfigure(2, weight=1)

        self.video_renamer_run_button_wrapper = ttk.Frame(self.video_renamer_run_control_frame, style='TFrame')
        self.video_renamer_run_button_wrapper.grid(row=0, column=1, sticky="")
        self.run_video_renamer_button = ttk.Button(self.video_renamer_run_button_wrapper, text="Run Video Renamer", command=self._run_video_renamer_script, style='TButton')
        self.run_video_renamer_button.pack(padx=5, pady=0)

        self.video_renamer_progress_wrapper = ttk.Frame(self.video_renamer_run_control_frame, style='TFrame')
        self.video_renamer_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.video_renamer_progress_bar = ttk.Progressbar(self.video_renamer_progress_wrapper, orient="horizontal", length=200, mode="indeterminate")
        self.video_renamer_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.video_renamer_progress_label = ttk.Label(self.video_renamer_progress_wrapper, text="", style='TLabel')
        self.video_renamer_progress_label.pack(side="right", padx=5)
        self.video_renamer_progress_wrapper.grid_remove()

        row_counter += 1

        # SECTION: Bynder Video Metadata Prep
        bynder_video_metadata_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        bynder_video_metadata_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_bynder_video_metadata = ttk.Frame(bynder_video_metadata_wrapper_frame, style='TFrame')
        header_sub_frame_bynder_video_metadata.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_bynder_video_metadata = ttk.Label(header_sub_frame_bynder_video_metadata, text="Bynder Video Metadata Export", style='Header.TLabel')
        header_label_bynder_video_metadata.pack(side="left", padx=(0, 5))
        info_label_bynder_video_metadata = ttk.Label(header_sub_frame_bynder_video_metadata, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_bynder_video_metadata, "Generates a semicolon-delimited CSV for Bynder import based on a spreadsheet and a folder of *renamed* video assets. Allows manual override for Wrike Link, Year, Sub-Initiative, and Location Type for the batch.", self.secondary_bg, self.text_color)  
        info_label_bynder_video_metadata.pack(side="left", anchor="center")

        bynder_video_metadata_frame = ttk.Frame(bynder_video_metadata_wrapper_frame, style='TFrame')
        bynder_video_metadata_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        bynder_video_metadata_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(bynder_video_metadata_frame, text="Spreadsheet:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(bynder_video_metadata_frame, textvariable=self.bynder_video_metadata_spreadsheet_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(bynder_video_metadata_frame, text="Browse File", command=lambda: self._browse_file(self.bynder_video_metadata_spreadsheet_path, "spreadsheet"), style='TButton').grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(bynder_video_metadata_frame, text="Folder of Renamed Assets (Videos):", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(bynder_video_metadata_frame, textvariable=self.bynder_video_metadata_folder_path, width=45, style='TEntry').grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(bynder_video_metadata_frame, text="Browse Folder", command=lambda: self._browse_folder(self.bynder_video_metadata_folder_path), style='TButton').grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(bynder_video_metadata_frame, text="Link to Wrike Project:", style='TLabel').grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(bynder_video_metadata_frame, textvariable=self.link_to_wrike_project_ui, width=45, style='TEntry').grid(row=2, column=1, padx=5, pady=5, sticky="ew", columnspan=2)
        Tooltip(bynder_video_metadata_frame.winfo_children()[-1], "Enter a Wrike project link to apply to all assets. If left blank, the script will attempt to use the value from the spreadsheet, if available.", self.secondary_bg, self.text_color)

        ttk.Label(bynder_video_metadata_frame, text="Year:", style='TLabel').grid(row=3, column=0, padx=5, pady=5, sticky="w")
        year_combobox = ttk.Combobox(bynder_video_metadata_frame, textvariable=self.year_ui, values=self.available_years, state="readonly", width=10, style='TCombobox')
        year_combobox.grid(row=3, column=1, padx=5, pady=5, sticky="w", columnspan=2)
        Tooltip(year_combobox, "Select the year to apply to all assets. If left blank, the script will attempt to use the value from the spreadsheet, if available.", self.secondary_bg, self.text_color)

        ttk.Label(bynder_video_metadata_frame, text="Sub-Initiative:", style='TLabel').grid(row=4, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(bynder_video_metadata_frame, textvariable=self.sub_initiative_ui, width=45, style='TEntry').grid(row=4, column=1, padx=5, pady=5, sticky="ew", columnspan=2)
        Tooltip(bynder_video_metadata_frame.winfo_children()[-1], "Enter the Sub-Initiative to apply to all assets. If left blank, the script will attempt to use the value from the spreadsheet, if available.", self.secondary_bg, self.text_color)

        ttk.Label(bynder_video_metadata_frame, text="Location Type:", style='TLabel').grid(row=5, column=0, padx=5, pady=5, sticky="w")
        location_type_combobox = ttk.Combobox(bynder_video_metadata_frame, textvariable=self.location_type_ui, values=self.location_type_options, state="readonly", width=30, style='TCombobox')
        location_type_combobox.grid(row=5, column=1, padx=5, pady=5, sticky="w", columnspan=2)
        Tooltip(location_type_combobox, "Select the Location Type to apply to all assets. If left blank, the script will attempt to use the value from the spreadsheet, if available.", self.secondary_bg, self.text_color)

        self.bynder_video_metadata_export_run_control_frame = ttk.Frame(bynder_video_metadata_frame, style='TFrame')
        self.bynder_video_metadata_export_run_control_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky="ew")
        self.bynder_video_metadata_export_run_control_frame.grid_columnconfigure(0, weight=1)
        self.bynder_video_metadata_export_run_control_frame.grid_columnconfigure(1, weight=0)
        self.bynder_video_metadata_export_run_control_frame.grid_columnconfigure(2, weight=1)

        self.bynder_video_metadata_export_run_button_wrapper = ttk.Frame(self.bynder_video_metadata_export_run_control_frame, style='TFrame')
        self.bynder_video_metadata_export_run_button_wrapper.grid(row=0, column=1, sticky="")
        self.run_bynder_video_metadata_export_button = ttk.Button(self.bynder_video_metadata_export_run_button_wrapper, text="Export Bynder Metadata CSV", command=self._run_bynder_video_metadata_export, style='TButton')
        self.run_bynder_video_metadata_export_button.pack(padx=5, pady=0)

        self.bynder_video_metadata_export_progress_wrapper = ttk.Frame(self.bynder_video_metadata_export_run_control_frame, style='TFrame')
        self.bynder_video_metadata_export_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.bynder_video_metadata_export_progress_bar = ttk.Progressbar(self.bynder_video_metadata_export_progress_wrapper, orient="horizontal", length=200, mode="indeterminate")
        self.bynder_video_metadata_export_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.bynder_video_metadata_export_progress_label = ttk.Label(self.bynder_video_metadata_export_progress_wrapper, text="", style='TLabel')
        self.bynder_video_metadata_export_progress_label.pack(side="right", padx=5)
        self.bynder_video_metadata_export_progress_wrapper.grid_remove()

        row_counter += 1

        # --- Activity Log ---
        self.log_wrapper_frame = ttk.Frame(self.master, style='SectionFrame.TFrame')
        self.log_wrapper_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")  

        self.log_header_frame = ttk.Frame(self.log_wrapper_frame, style='TFrame')
        self.log_header_frame.pack(fill="x", padx=5, pady=2, side="top")  
        
        log_title_label = ttk.Label(self.log_header_frame, text="Activity Log", font=self.header_font, foreground=self.header_text_color, background=self.secondary_bg)
        log_title_label.pack(side="left", padx=(0, 5))  
        
        self.toggle_log_button = ttk.Button(self.log_header_frame, text="▼", command=self._toggle_log_size, width=2, style='TButton')
        self.toggle_log_button.pack(side="right")


        self.log_text = scrolledtext.ScrolledText(self.log_wrapper_frame, width=90, height=15,  
                                                  font=self.log_font, state='disabled',
                                                  bg=self.log_bg, fg=self.log_text_color,
                                                  insertbackground=self.log_text_color,  
                                                  selectbackground=self.accent_color,  
                                                  selectforeground=self.RF_WHITE_BASE,  
                                                  relief="solid", borderwidth=1)
        self.log_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)  

        if not self.log_expanded:  
            self.log_text.pack_forget()  
            self.toggle_log_button.config(text="▲")  
            self.master.grid_rowconfigure(2, weight=0)  
            self.log_wrapper_frame.config(height=50)  
        

if __name__ == "__main__":
    root = tk.Tk()
    app = AdVideoApp(root)

    creator_frame = ttk.Frame(root, style='TFrame')
    creator_frame.grid(row=3, column=0, sticky="se", padx=10, pady=5)
    creator_label = ttk.Label(creator_frame, text="Created By: Zachary Eisele", font=("Arial", 8), foreground="#888888", background=root.cget('bg'))
    creator_label.pack(side="right", anchor="se")

    root.mainloop()
