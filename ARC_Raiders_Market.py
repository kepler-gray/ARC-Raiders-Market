import customtkinter as ctk
import keyboard
import json
import win32gui
import os
import re
import pyautogui
from difflib import get_close_matches
import time 
import threading
import tkinter.messagebox as tk_messagebox
import sys 

# Windows-specific libraries for shortcut creation and tray icon
import win32com.client
import shutil
from pystray import Icon as TrayIcon, Menu as Menu, MenuItem as Item
from PIL import Image

# --- CONFIGURATION ---
HOTKEY = 'ctrl + f'
HOTKEY_ESC = 'esc'
TARGET_WINDOW_KEYWORD = "ARC Raiders" 
APP_TITLE = "ARC Raiders Market"
JSON_FILENAME = "ARC_Raiders_Market_Database.json" 
ICON_FILENAME = "app_icon.ico" 

# Rarity to Hex Color Map
RARITY_COLORS = {
    "COMMON": "#8C8C8C",    
    "UNCOMMON": "#26BF57",  
    "RARE": "#00A8F2",      
    "EPIC": "#DF40AA",      
    "LEGENDARY": "#FFC600"  
}

# --- DEFAULT EMBEDDED DATA (Fallback for compiled EXE) ---
DEFAULT_ITEM_DB = {
    "Matriarch Reactor": {"price": 13000, "rarity": "Legendary"},
    "Magnetron": {"price": 6000, "rarity": "Epic"},
    "Microscope": {"price": 3000, "rarity": "Rare"},
    "Duct Tape": {"price": 300, "rarity": "Uncommon"},
    "Chemicals": {"price": 50, "rarity": "Common"},
    "Crystallized Hydroxide": {"price": 1050, "rarity": "Rare"},
    "Infrared Lens": {"price": 800, "rarity": "Uncommon"},
    "Quantum Entangler": {"price": 25000, "rarity": "Legendary"},
    "Fusion Coil": {"price": 9500, "rarity": "Epic"},
    "Nanotube Fiber": {"price": 450, "rarity": "Uncommon"},
    "Scrap Metal": {"price": 10, "rarity": "Common"}
}

# Global Data Variables
ITEM_DB, ITEM_NAMES, ITEM_COLORS, LOWER_TO_ORIGINAL_MAP = {}, [], {}, {}

# --- DATA HANDLING (FIXED for EXE location) ---
def load_data():
    """Tries to load data from external JSON (next to EXE/script) 
       and falls back to embedded default data if file is missing or corrupted."""
    
    data_to_use = DEFAULT_ITEM_DB
    
    # 1. Determine the correct directory path (same place as EXE/script)
    if getattr(sys, 'frozen', False):
        # Compiled EXE: use the directory of the executable
        base_dir = os.path.dirname(sys.executable)
    else:
        # Running from source: use the directory of the script file
        base_dir = os.path.dirname(os.path.abspath(__file__))

    json_path = os.path.join(base_dir, JSON_FILENAME)
    
    # 2. Try to load the external data
    try:
        with open(json_path, 'r') as f:
            data_to_use = json.load(f)
        print(f"Data loaded successfully from external file: {json_path}")

    except FileNotFoundError:
        print(f"File {JSON_FILENAME} not found at {json_path}. Using default embedded data.")
        
        # If running from source and file is missing, create it for user convenience
        if not getattr(sys, 'frozen', False):
            try:
                with open(json_path, 'w') as f:
                    json.dump(DEFAULT_ITEM_DB, f, indent=4)
                print(f"Created new {JSON_FILENAME} in source directory. Please update prices there.")
            except Exception as e:
                print(f"Error creating JSON file in source dir: {e}.")
        
    except json.JSONDecodeError:
        print(f"Error reading {JSON_FILENAME}. It might be corrupted. Using default embedded data.")
        
    except Exception as e:
        print(f"An unexpected error occurred during file operation: {e}. Using default embedded data.")


    # 3. Process the chosen data
    item_db = data_to_use
    lower_item_names = []
    item_colors = {}
    lower_to_original_map = {}
 
    for name, details in item_db.items():
        lower_name = name.lower()
        
        lower_item_names.append(lower_name)
        lower_to_original_map[lower_name] = name 
   
        rarity_key = details.get('rarity', 'COMMON').upper()
        item_colors[name] = RARITY_COLORS.get(rarity_key, RARITY_COLORS['COMMON'])
        
    return item_db, lower_item_names, item_colors, lower_to_original_map

# Global Data Variables initialized
ITEM_DB, ITEM_NAMES, ITEM_COLORS, LOWER_TO_ORIGINAL_MAP = load_data()

# --- STARTUP SHORTCUT LOGIC ---

FLAG_FILE_NAME = ".first_run_flag" 

def get_app_data_path(app_name):
    """Returns the path to the app's roaming AppData folder, creating it if necessary."""
    app_data_dir = os.path.join(os.getenv('APPDATA'), app_name)
    os.makedirs(app_data_dir, exist_ok=True)
    return app_data_dir

# Calculate the full path for the flag file
APP_DATA_DIR = get_app_data_path(APP_TITLE)
FULL_FLAG_PATH = os.path.join(APP_DATA_DIR, FLAG_FILE_NAME)


def create_startup_shortcut():
    """Prompts the user and creates a shortcut in the Windows Startup folder if confirmed."""
    
    if os.path.exists(FULL_FLAG_PATH):
        return

    should_create = tk_messagebox.askyesno(
        title="Startup Shortcut",
        message="Would you like ARC Raiders Market to start automatically in the background when Windows launches?"
    )

    try:
        with open(FULL_FLAG_PATH, 'w') as f:
            f.write("run=true")
    except Exception as e:
        print(f"Error writing first-run flag to AppData: {e}")

    if should_create:
        try:
            exe_path = os.path.realpath(__file__)
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            
            startup_folder = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
            shortcut_path = os.path.join(startup_folder, f"{APP_TITLE}.lnk")

            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = exe_path
            shortcut.WorkingDirectory = os.path.dirname(exe_path)
            shortcut.IconLocation = exe_path
            shortcut.save()
            print(f"Startup shortcut created at: {shortcut_path}")
        except Exception as e:
            print(f"Failed to create startup shortcut: {e}")
            tk_messagebox.showerror(title="Startup Error", message="Could not create the startup shortcut. You may need to run the application as administrator once.")

# --- TRAY ICON LOGIC ---

class AppTrayIcon:
    """Manages the system tray icon and its actions."""
    def __init__(self, app_instance):
        self.app = app_instance
        self.icon = self._create_icon()
        self.thread = threading.Thread(target=self._run_icon, daemon=True)

    def _create_icon(self):
        """Loads the icon file and creates the tray menu."""
        
        try:
            icon_image = Image.open(ICON_FILENAME)
        except Exception:
            print(f"Warning: Could not load icon file '{ICON_FILENAME}'. Using default white square.")
            icon_image = Image.new('RGB', (64, 64), color='white')
        
        menu = Menu(
            Item(f"{APP_TITLE} (Ctrl+F)", self.noop),
            Menu.SEPARATOR,
            Item("Show", self.show_app),
            Item("Exit", self.exit_app)
        )
        
        return TrayIcon(name=APP_TITLE, icon=icon_image, title=APP_TITLE, menu=menu)

    def _run_icon(self):
        """Starts the tray icon listener."""
        self.icon.run()

    def start(self):
        """Starts the icon thread."""
        self.thread.start()

    def stop(self):
        """Stops the icon and joins the thread."""
        self.icon.stop()
        
    def noop(self, icon, item):
        """A placeholder for non-actionable menu items."""
        pass

    def show_app(self, icon, item):
        """Forces the Tkinter app to show itself (via the main thread)."""
        self.app.after(0, self.app.show_overlay)

    def exit_app(self, icon, item):
        """Handles graceful exit of the entire application (via the main thread)."""
        self.app.after(0, self.app.quit_app)


# --- GUI CLASS ---
class SearchOverlay(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 1. Window Setup
        self.title("ARC Raiders Market")
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.resizable(False, False)
        self.configure(fg_color="#121216")
        
        # --- FIX RE-APPLIED: Override the standard 'X' close button behavior to minimize to tray ---
        # This protocol intercepts the window closing event (WM_DELETE_WINDOW) 
        # and runs our custom function (minimize_to_tray) instead of quitting.
        self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray) 
        # ------------------------------------------------------------------------------------------

        # Set the custom icon for the Tkinter window
        try:
            self.iconbitmap(ICON_FILENAME)
        except Exception as e:
            print(f"Warning: Failed to set window icon: {e}")

        
        # --- FONT SELECTION ---
        self.main_font = "Roboto Medium" 
        
 
        # 2. Calculate Center
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 400
        window_height = 200 
        x_pos = (screen_width // 2) - (window_width // 2)
        y_pos = (screen_height // 3)
        
        self.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")

     
        # 3. UI Layout
        self.grid_columnconfigure(0, weight=1)
        
        # Title
        self.lbl_title = ctk.CTkLabel(
            self, 
            text="ARC Raiders Sell Prices", 
            font=(self.main_font, 16),
            text_color="#555555"
        )
        self.lbl_title.pack(pady=(10, 0))

        # Input
        self.entry = ctk.CTkEntry(
            self, 
            placeholder_text="...", 
            width=500, 
            height=40,
            font=(self.main_font, 20),
            border_color="#333333",
            fg_color="#1e1e1e"
        )
        self.entry.pack(pady=(5, 21))
        self.entry.bind("<KeyRelease>", self.update_results)
        self.entry.bind("<Return>", self.hide_overlay)

        # Result Label
        self.result_label = ctk.CTkLabel(
            self, 
            text="Ready", 
            font=(self.main_font, 24, "bold"),
            text_color="#ffcc00"
        )
        self.result_label.pack(pady=(0, 10))
        
        # Done Button
        self.done_button = ctk.CTkButton(
            self, 
            text="Done", 
            command=self.done_clicked,
            width=100,
            height=30,
            font=(self.main_font, 16, "bold"),
            fg_color="#333333",
            hover_color="#555555"
        )
        self.done_button.pack(pady=(5, 15)) 
        
  
        # ESCAPE KEY TRACKER
        self.esc_listener = None 

        # Start hidden
        self.withdraw()
        
        # Reference to the tray icon manager
        self.tray_icon_manager = None

    def minimize_to_tray(self):
        """Called when the user clicks the 'X' button to hide the app instead of closing."""
        print("Window close requested, minimizing to tray...")
        self.hide_overlay()

    def quit_app(self):
        """Cleans up listeners and safely exits the application."""
        print("Quitting application...")
 
        # 1. Stop hotkey listener (essential for clean exit)
        try:
            keyboard.unhook_all()
        except Exception:
            pass
            
        # 2. Stop tray icon thread
        if self.tray_icon_manager:
            self.tray_icon_manager.stop()
            
        # 3. Destroy main window and exit mainloop
        self.destroy()

    def update_results(self, event):
        query = self.entry.get()
        
        if len(query) < 3:
            self.result_label.configure(text="...", text_color="#777777")
            return
        
        lower_query = query.lower()
        matches = get_close_matches(lower_query, ITEM_NAMES, n=1, cutoff=0.5)
        
        if matches:
            matched_lower_name = matches[0]
            name = LOWER_TO_ORIGINAL_MAP[matched_lower_name]
            details = ITEM_DB[name]
        
            price = details['price']
            color = ITEM_COLORS[name]
            
            formatted_price = "{:,}".format(price)
            self.result_label.configure(text=f"{name}: ${formatted_price}", text_color=color)
        else:
            self.result_label.configure(text="No match", text_color="#555555")

    # --- THREADED CLICK FUNCTIONS (Omitted for brevity, kept same as original) ---
  
    def _threaded_focus_click(self, x, y):
        start_time = time.time()
        print(f"\n[FOCUS THREAD] Starting pyautogui click (Input box) at {start_time}")
        try:
            pyautogui.click(x=x, y=y)
        except Exception as e:
            print(f"[FOCUS THREAD] Error during click: {e}")
        end_time = time.time()
        print(f"[FOCUS THREAD] Click finished at {end_time}. Duration: {end_time - start_time:.4f} seconds.")
        self.after(0, self.entry.focus_force)

    def click_and_focus(self):
        entry_x_center_rel = self.entry.winfo_x() + (self.entry.winfo_width() // 2)
        entry_y_center_rel = self.entry.winfo_y() + (self.entry.winfo_height() // 2)
        absolute_x = self.winfo_x() + entry_x_center_rel
        absolute_y = self.winfo_y() + entry_y_center_rel
        
        t = threading.Thread(target=self._threaded_focus_click, args=(absolute_x, absolute_y,))
        t.start()


    def _threaded_center_click(self):
        start_time = time.time()
        print(f"\n[CENTER THREAD] Starting pyautogui click (Center screen) at {start_time}")
        try:
            screen_width, screen_height = pyautogui.size()
            center_x = screen_width // 2
            center_y = screen_height // 2
            
            pyautogui.click(center_x, center_y)
        except Exception as e:
            print(f"[CENTER THREAD] Error forcing center click: {e}")
            
        end_time = time.time()
        print(f"[CENTER THREAD] Click finished at {end_time}. Duration: {end_time - start_time:.4f} seconds.")

    def _force_center_click(self):
        t = threading.Thread(target=self._threaded_center_click)
        t.start()
        
    # --- WINDOW CONTROL FUNCTIONS ---

    def done_clicked(self):
        self.hide_overlay()
        
        target_keyword = TARGET_WINDOW_KEYWORD.upper()
        
        start_time = time.time() 
        print(f"\n[DONE] Starting SetForegroundWindow operation at {start_time}")

        def callback(hwnd, extra):
            window_title = win32gui.GetWindowText(hwnd)
            if target_keyword in window_title.upper():
                win32gui.SetForegroundWindow(hwnd)
                raise Exception("Window found and focused")

        try:
            win32gui.EnumWindows(callback, None)
        except Exception as e:
            if str(e) != "Window found and focused":
                 print(f"Error while trying to restore focus to game: {e}")
        
        end_time = time.time() 
        print(f"[DONE] SetForegroundWindow finished at {end_time}. Duration: {end_time - start_time:.4f} seconds.")

        self.after(50, self._force_center_click)

    def show_overlay(self):
        self.deiconify()
        self.attributes('-topmost', True)

        try:
            self.lift() 
        except Exception:
            self.lift() 
            
        self.entry.delete(0, 'end')
 
        self.result_label.configure(text="Waiting for input...", text_color="#777777")
        
        self.after(50, self.click_and_focus)

        if self.esc_listener is None:
            self.esc_listener = keyboard.add_hotkey(HOTKEY_ESC, lambda: self.after(0, self.hide_overlay))

    def hide_overlay(self, event=None):
        self.withdraw()
        
        if self.esc_listener:
            try:
                keyboard.remove_hotkey(self.esc_listener)
                self.esc_listener = None
            except Exception:
                pass

# --- SYSTEM LOGIC ---

def get_active_window_title():
    """Returns the title of the window currently in the foreground."""
    try:
        hwnd = win32gui.GetForegroundWindow()
        return win32gui.GetWindowText(hwnd)
    except Exception:
       return ""

def clean_string(text):
    """Removes non-alphanumeric characters."""
    return re.sub(r'[^a-zA-Z0-9]', '', text).upper()

def on_hotkey():
    raw_title = get_active_window_title()
    cleaned_title = clean_string(raw_title)
    target_clean = clean_string(TARGET_WINDOW_KEYWORD)
    
    # Check if the target game window is active
    if cleaned_title.startswith(target_clean):
        if len(target_clean) >= 3: 
            app.after(0, app.show_overlay)
    else:
        # Ignore if not the target window
        pass

# Start the hotkey listener
keyboard.add_hotkey(HOTKEY, on_hotkey, suppress=False)

print(f"Started. Listening for '{HOTKEY}'.")


if __name__ == "__main__":
    
    # --- 1. Handle Startup Shortcut ---
    create_startup_shortcut()
    
    # --- 2. Initialize App and Tray Icon ---
    app = SearchOverlay()
    
    tray_manager = AppTrayIcon(app)
    app.tray_icon_manager = tray_manager # Give the app a reference to the tray manager
    tray_manager.start() # Start the tray icon thread
    
    # --- 3. Start Main Loop ---
    try:
        app.mainloop()

    except Exception as e:
        print("-" * 50)
        print("🚨 FATAL ERROR: The application crashed.")
        print(f"Details: {e}")
        print("-" * 50)
        
        if not getattr(sys, 'frozen', False):
            input("Press Enter to close...")

    # Final cleanup outside the main loop
    if app.tray_icon_manager:
        app.tray_icon_manager.stop()
        

    print("Application terminated cleanly.")
