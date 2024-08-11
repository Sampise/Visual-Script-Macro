import customtkinter as ctk
import win32con
import win32gui
import json
from pyautogui import keyDown, keyUp
from pynput import keyboard


WHITE = "#FFFFFF"
VERY_DARK_GRAY  = "#242424"
ALMOST_BLACK  = "#181818"
DARK_GRAY = "#2b2b2b"
MEDIUM_DARK_GRAY = "#333333"

def load_data():
    with open("macros.json", 'r') as file:
        return json.load(file)


class TitleBar(ctk.CTkFrame):
    def __init__(self, parent, on_close, title: str, minimize: bool):
        super().__init__(parent, bg_color=VERY_DARK_GRAY , fg_color=VERY_DARK_GRAY )
        self.root = parent
        self.on_close = on_close
        self.root.overrideredirect(True)
        
        self.nav_title = ctk.CTkLabel(self, text=title, fg_color=VERY_DARK_GRAY , bg_color=VERY_DARK_GRAY )
        self.nav_title.pack(side="left", padx=10)
        
        self.close_button = ctk.CTkButton(self, text='✕', cursor="hand2", corner_radius=0, fg_color=VERY_DARK_GRAY , hover_color=ALMOST_BLACK , width=40, command=self.close_window)
        self.close_button.pack(side="right")
        
        if minimize:
            self.minimize_button = ctk.CTkButton(self, text='—', cursor="hand2", corner_radius=0, fg_color=VERY_DARK_GRAY , hover_color=ALMOST_BLACK , width=40, command=self.minimize_window)
            self.minimize_button.pack(side="right")
        
        self.bind("<ButtonPress-1>", self.oldxyset)
        self.bind("<B1-Motion>", self.move)
        
        self.nav_title.bind("<ButtonPress-1>", self.oldxyset_label)
        self.nav_title.bind("<B1-Motion>", self.move)

    def oldxyset(self, event):
        self.oldx = event.x
        self.oldy = event.y

    def oldxyset_label(self, event):
        self.oldx = event.x + self.nav_title.winfo_x()
        self.oldy = event.y + self.nav_title.winfo_y()
        
    def move(self, event):
        new_x = event.x_root - self.oldx
        new_y = event.y_root - self.oldy
        self.root.geometry(f"+{new_x}+{new_y}")
    
    def close_window(self):
        self.on_close()
        
    def minimize_window(self):
        window_id = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(window_id, win32con.SW_MINIMIZE)

        
class SettingsPopup(ctk.CTkToplevel):
    def __init__(self, parent, mainWindow, currentstart, currentstop):
        super().__init__(parent)
        self.mainWindow = mainWindow
        self.title("Settings")
        self.geometry("300x200")
        self.resizable(False, False)
        self.overrideredirect(True)
        self.transient(parent)
        
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f"300x200+{int(screen_width / 2 - 150)}+{int(screen_height / 2 - 100)}")
        self.grab_set()

        self.key_for_button1 = currentstart
        self.key_for_button2 = currentstop
        self.capture_input = None
        
        TitleBar(self, title="Settings", on_close=self.on_close, minimize=False).pack(fill="x")
        
        button_frame = ctk.CTkFrame(self, corner_radius=5, fg_color=DARK_GRAY)
        button_frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        self.button1 = ctk.CTkButton(button_frame, text=f"Start - {self.key_for_button1}", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=lambda: self.capture_key_input("button1"))
        self.button1.pack(pady=5, padx=5, side="top")
        
        self.button2 = ctk.CTkButton(button_frame, text=f"Stop - {self.key_for_button2}", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=lambda: self.capture_key_input("button2"))
        self.button2.pack(pady=5, padx=5, side="top")
        
        self.label = ctk.CTkLabel(button_frame, text="0-9 | F1-F12", text_color="#6d6d6d")
        self.label.pack(pady=0, padx=5, side="top")
        
        self.apply_button = ctk.CTkButton(button_frame, text="Apply", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.apply_changes)
        self.apply_button.pack(pady=10, padx=5, side="top")

        self.bind("<Key>", self.on_key_pressed)

    def on_key_pressed(self, event):
        allowed_keys = [str(i) for i in range(10)] + [f"F{i}" for i in range(1, 13)]
        
        if event.keysym in allowed_keys:
            if self.capture_input == "button1":
                self.key_for_button1 = event.keysym
                self.update_button_text("button1", self.key_for_button1)
                self.capture_input = None
            elif self.capture_input == "button2":
                self.key_for_button2 = event.keysym
                self.update_button_text("button2", self.key_for_button2)
                self.capture_input = None
        else:
            self.label.configure(text_color="#ff0000")
            self.after(200, lambda: self.label.configure(text_color="#6d6d6d"))
            
    def capture_key_input(self, button):
        self.capture_input = button

    def update_button_text(self, button, new_text):
        if button == "button1":
            self.button1.configure(text=f"Start - {new_text}")
        elif button == "button2":
            self.button2.configure(text=f"Stop - {new_text}")

    def apply_changes(self):
        MacroApp.update_keybind(self.mainWindow, new_start_key=self.key_for_button1, new_stop_key=self.key_for_button2)
        MacroApp.update_button_text(self.mainWindow)
        MacroApp.restart_listener(self.mainWindow)
        self.destroy()

    def on_close(self):
        self.destroy()

class CustomFrameItem(ctk.CTkFrame):
    def __init__(self, parent, label_text, editor_instance, action_data, characters=None):
        from PIL import Image
        super().__init__(parent, fg_color=MEDIUM_DARK_GRAY)
        self.editor_instance = editor_instance 
        
        self.action_data = action_data
        self.icon = ctk.CTkLabel(self, text="", bg_color=VERY_DARK_GRAY , fg_color=VERY_DARK_GRAY , corner_radius=0, width=30, height=30)

        if characters is None:
            if not hasattr(self.__class__, 'preloaded_image'):
                def resource_path(relative_path):
                    import os
                    import sys
                    try:
                        base_path = sys._MEIPASS
                    except Exception:
                        base_path = os.path.abspath(".")

                    return os.path.join(base_path, relative_path)
                self.__class__.preloaded_image = Image.open(resource_path("zzz.png"))
            self.ctk_image = ctk.CTkImage(self.__class__.preloaded_image, size=(30, 30))
            self.icon.configure(image=self.ctk_image)
        else:
            self.icon.configure(text=characters, font=("Arial", 15, "bold"))
            
        self.icon.pack(side="left", fill="x")
        self.spacer = ctk.CTkFrame(self, width=10, height=30, fg_color=VERY_DARK_GRAY , bg_color=VERY_DARK_GRAY ).pack(side="left")
        
        self.label = ctk.CTkLabel(self, text=label_text, corner_radius=0, height=30, width=120, wraplength=120, anchor="w", bg_color=VERY_DARK_GRAY , font=("Arial", 16))
        self.label.pack(side="left", padx=(0, 0), pady=0, fill="x")
        
        self.bind("<Button-3>", self.delete_item)
        self.icon.bind("<Button-3>", self.delete_item)
        self.label.bind("<Button-3>", self.delete_item)
        
    def delete_item(self, event):
        self.editor_instance.items.remove(self.action_data)
        self.pack_forget() 

class ActionPopup(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.editor_popup_instance = parent
        self.title("Create Macro")
        self.resizable(False, False)
        self.overrideredirect(True)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        window_width, window_height = 600, 500
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f"{window_width}x{window_height}+{(screen_width - window_width) // 2}+{(screen_height - window_height) // 2}")
        self.grab_set()
        
        TitleBar(self, title="Add action", on_close=self.on_close,minimize=False).pack(fill="x")


        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=DARK_GRAY)
        self.main_frame.pack(expand=True, fill='both', padx=10, pady=10)

        self.first_section = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color=DARK_GRAY)
        self.first_section.grid(row=0, column=0, sticky='news')

        self.first_button_frame = ctk.CTkFrame(self.first_section, fg_color=VERY_DARK_GRAY )
        self.first_button_frame.pack(expand=True, side="left")
        self.first_button_label = ctk.CTkLabel(self.first_button_frame, text="First key:", font=("Arial", 16))
        self.first_button_label.pack(pady=(10, 0))
        self.first_button_input = ctk.CTkButton(self.first_button_frame, text="(required)", height=35, width=200, border_color=VERY_DARK_GRAY , fg_color=MEDIUM_DARK_GRAY, command=lambda: self.capture_key_input("macro_name_input1"))
        self.first_button_input.pack(pady=(10, 10), padx=10)

        self.second_button_frame = ctk.CTkFrame(self.first_section, fg_color=VERY_DARK_GRAY )
        self.second_button_frame.pack(expand=True, side="right")
        self.second_button_label = ctk.CTkLabel(self.second_button_frame, text="Second key:", font=("Arial", 16))
        self.second_button_label.pack(pady=(10, 0))
        self.second_button_input = ctk.CTkButton(self.second_button_frame, text="(optional)", height=35, width=200, border_color=VERY_DARK_GRAY , fg_color=MEDIUM_DARK_GRAY, command=lambda: self.capture_key_input("macro_name_input2"))
        self.second_button_input.pack(pady=(10, 10), padx=10)

        self.second_section = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color=DARK_GRAY)
        self.second_section.grid(row=1, column=0, sticky='news')

        self.action_duration_frame = ctk.CTkFrame(self.second_section, fg_color=VERY_DARK_GRAY )
        self.action_duration_frame.pack(expand=True)
        self.action_duration_label = ctk.CTkLabel(self.action_duration_frame, text="Action Duration:", font=("Arial", 16))
        self.action_duration_label.pack(pady=(10, 0))
        self.action_duration_input = ctk.CTkEntry(self.action_duration_frame, placeholder_text="(Seconds)", height=35, width=200, border_color=VERY_DARK_GRAY , fg_color=MEDIUM_DARK_GRAY, validate="key", validatecommand=(self.register(self.validate_input), "%P"))
        self.action_duration_input.pack(pady=(10, 10), padx=10)

        self.third_section = ctk.CTkFrame(self.main_frame, corner_radius=0, fg_color=DARK_GRAY)
        self.third_section.grid(row=2, column=0, sticky='news')

        self.confirm_button = ctk.CTkButton(self.third_section, height=28, width=112, text="Confirm", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.send_properties, state="disabled")
        self.confirm_button.pack(side="bottom", pady=5)

        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.capture_input = None  
        self.key_for_button1 = None
        self.key_for_button2 = None

        self.bind("<Key>", self.on_key_pressed)
        self.load_data = load_data
    
    
    def send_properties(self):
        if self.key_for_button2: 
            self.editor_popup_instance.add_action(time=self.action_duration_input.get(),action_one=self.key_for_button1, action_two=self.key_for_button2)
        else:
            self.editor_popup_instance.add_action(time=self.action_duration_input.get(),action_one=self.key_for_button1)
        self.on_close()
    
    def validate_input(self, input_value):
        import re
        if input_value == "" or re.match(r'^(?=\d*(?:\.\d*)?$)(?=.{1,6}$)\d*\.?\d*$', input_value):
            return True
        return False
    
    def on_key_pressed(self, event):
        allowed_keys = [str(i) for i in range(10)] + [chr(i) for i in range(ord('a'), ord('z') + 1)] + [chr(i) for i in range(ord('A'), ord('Z') + 1)]

        key = event.keysym.lower()
        self.check_entries()
        
        if key in allowed_keys:
            if self.capture_input == "macro_name_input1":
                self.key_for_button1 = key
                self.update_button_text("macro_name_input1", self.key_for_button1)
                self.capture_input = None 
            elif self.capture_input == "macro_name_input2":
                self.key_for_button2 = key
                self.update_button_text("macro_name_input2", self.key_for_button2)
                self.capture_input = None  

    def capture_key_input(self, button):
        self.capture_input = button
        
    def update_button_text(self, button, new_text):
        if button == "macro_name_input1":
            self.first_button_input.configure(text=f"{new_text}")
        elif button == "macro_name_input2":
            self.second_button_input.configure(text=f"{new_text}")
            
    def check_entries(self, *args):
        if self.key_for_button1 and self.action_duration_input.get():
            self.confirm_button.configure(state='normal') 
        else:
            self.confirm_button.configure(state='disabled') 
    
    def on_close(self):
        self.destroy()
        
class SleepPopup(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.editor_popup_instance = parent
        self.title("Create Macro")
        self.resizable(False, False)
        self.overrideredirect(True)
        self.transient(parent)
        
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width, window_height = 300, 300
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.grab_set()
        
        TitleBar(self, title="Add sleep", on_close=self.on_close,minimize=False).pack(fill="x")  
        
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=DARK_GRAY)
        self.main_frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        self.sleep_duration_frame = ctk.CTkFrame(self.main_frame, fg_color=VERY_DARK_GRAY )
        self.sleep_duration_frame.pack(expand=True)
        self.sleep_duration_label = ctk.CTkLabel(self.sleep_duration_frame, text="Action Duration:", font=("Arial", 16))
        self.sleep_duration_label.pack(pady=(10, 0))
        self.sleep_duration_input = ctk.CTkEntry(self.sleep_duration_frame, placeholder_text="(Seconds)", height=35, width=200, border_color=VERY_DARK_GRAY , fg_color=MEDIUM_DARK_GRAY, validate="key", validatecommand=(self.register(self.validate_input), "%P"))
        self.sleep_duration_input.pack(pady=(10, 10), padx=10)
        
        self.confirm_button = ctk.CTkButton(self.main_frame, height=28, width=112, text="Confirm", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.send_properties, state="disabled")
        self.confirm_button.pack(side="bottom", pady=5)
        
        self.bind("<Key>", self.check_entries)
    def check_entries(self, *args):
        if self.sleep_duration_input.get():
            self.confirm_button.configure(state='normal') 
        else:
            self.confirm_button.configure(state='disabled') 
    
    def send_properties(self):
        self.editor_popup_instance.add_action(time=self.sleep_duration_input.get())
        self.on_close() 
          
    def validate_input(self, input_value):
        import re
        if input_value == "" or re.match(r'^(?=\d*(?:\.\d*)?$)(?=.{1,6}$)\d*\.?\d*$', input_value):
            return True
        return False
        
    def on_close(self):
        self.destroy()
            
class EditorPopup(ctk.CTkToplevel):
    def __init__(self, parent,mainWindow,existing_macro=False):
        self.mainWindow = mainWindow
        super().__init__(parent)
        self.title("Macro Editor")
        self.resizable(False, False)
        self.overrideredirect(True)
        self.transient(parent)
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width, window_height = 600, 500
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.grab_set() 
        TitleBar(self, title="Macro Editor", on_close=self.on_close,minimize=False).pack(fill="x")
        self.items = []
        self.mainFrame = ctk.CTkFrame(self, corner_radius=0, fg_color=DARK_GRAY)
        self.mainFrame.pack(expand=True, fill='both', padx=10, pady=10)
        self.mainFrame.grid_columnconfigure(0, weight=1)
        self.mainFrame.grid_columnconfigure(1, weight=3)
        self.mainFrame.grid_rowconfigure(0, weight=1)
        
        self.leftFrame = ctk.CTkFrame(self.mainFrame)
        self.leftFrame.grid(row=0, column=0, sticky='news')

        self.scrollable_frame = ctk.CTkScrollableFrame(self.leftFrame, corner_radius=4, label_text="Macro Actions:", label_fg_color=VERY_DARK_GRAY , fg_color=MEDIUM_DARK_GRAY)
        self.scrollable_frame.pack(side="top", fill="both", expand=True)
    
        self.add_minus_one()
        self.create_plus_button()
        self.create_right_frame()

        if existing_macro:
            self.populate_existing_macro(existing_macro)
        
    def add_minus_one(self):
        def add_minus_one():
            self.repetitionsInput.delete(0, ctk.END)
            self.repetitionsInput.insert(0, "-1")
            self.check_entries() 
        return add_minus_one
    
    def create_plus_button(self):
        testframe2 = ctk.CTkFrame(self.leftFrame, fg_color=DARK_GRAY)
        testframe2.pack(side="top", fill="x")
        
        def plus_button_function():
            self.plus_button.pack_forget()
            self.action_choice_button.pack(side="left", padx=(10,0), pady=5)
            self.sleep_choice_button.pack(side="right", padx=(0,10), pady=5)
            
        self.plus_button = ctk.CTkButton(testframe2, text="+", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , width=30, command=plus_button_function)
        self.plus_button.pack(side="top", padx=10, pady=5)
        
        self.action_choice_button = ctk.CTkButton(testframe2, text="Action", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , width=90, command=self.create_action_popup)
        self.sleep_choice_button = ctk.CTkButton(testframe2, text="Sleep", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , width=90, command=self.create_sleep_popup)
    
    def create_action_popup(self):
        self.withdraw()
        self.action_popup = ActionPopup(self)
        self.wait_window(self.action_popup) 
        self.deiconify() 
        
    def create_sleep_popup(self):
        self.withdraw()
        self.sleep_popup = SleepPopup(self)        
        self.wait_window(self.sleep_popup) 
        self.deiconify() 
        
    def create_right_frame(self):
        self.rightFrame = ctk.CTkFrame(self.mainFrame, fg_color=DARK_GRAY)
        self.rightFrame.grid(row=0, column=1, sticky='news')
        
        self.nameFrame = ctk.CTkFrame(self.rightFrame, fg_color=VERY_DARK_GRAY )
        self.nameFrame.pack(pady=(50,55), padx=75, fill="both", expand=False)
        nameLabel= ctk.CTkLabel(self.nameFrame, text="Macro name:", font=("Arial", 16))
        nameLabel.pack(pady=(10,0))
        self.nameInput= ctk.CTkEntry(self.nameFrame, placeholder_text="Name", height=35, width=240, border_color=VERY_DARK_GRAY , fg_color=MEDIUM_DARK_GRAY)
        self.nameInput.pack(pady=(10,10), padx=10)
        
        repetitionsFrame = ctk.CTkFrame(self.rightFrame, fg_color=VERY_DARK_GRAY )
        repetitionsFrame.pack(pady=(15,20), padx=75, expand=False, fill="both")

        repetitionsLabel = ctk.CTkLabel(repetitionsFrame, text="Number of repetitions:", font=("Arial", 16))
        repetitionsLabel.pack(pady=(10,0))

        inputButtonFrame = ctk.CTkFrame(repetitionsFrame, fg_color=VERY_DARK_GRAY )
        inputButtonFrame.pack(pady=10, padx=0, fill="x")

        self.repetitionsInput = ctk.CTkEntry(inputButtonFrame, placeholder_text="Repetitions", height=35, width=165, border_color=VERY_DARK_GRAY , fg_color=MEDIUM_DARK_GRAY, validate="key", validatecommand=(self.register(self.validate_input), "%P"))
        self.repetitionsInput.pack(side="left", padx=(10,0))

        infiniteButton = ctk.CTkButton(inputButtonFrame, text="∞", height=30, width=30, fg_color=MEDIUM_DARK_GRAY, hover_color=DARK_GRAY, command=self.add_minus_one())
        infiniteButton.pack(side="right", padx=(0,10)) 
        
        self.confirmButton = ctk.CTkButton(self.rightFrame, height=28, width=112, text="Confirm", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.confirm, state="disabled")
        self.confirmButton.pack(side="bottom", padx=50, pady=5)
        
        self.nameInput.bind("<KeyRelease>", self.check_entries)
        self.repetitionsInput.bind("<KeyRelease>", self.check_entries)

    def check_entries(self, *args):
        if self.nameInput.get() and self.repetitionsInput.get() and len(self.scrollable_frame.winfo_children()) > 0:
            self.confirmButton.configure(state='normal') 
        else:
            self.confirmButton.configure(state='disabled') 

    def confirm(self):
        new_macro = {
            "name": self.nameInput.get(),
            "repeatFor": int(self.repetitionsInput.get()),
            "actions": self.items
        }
        current_json = load_data()
        
        if 'macros' in current_json and isinstance(current_json['macros'], list):
            macro_exists = False
            for idx, macro in enumerate(current_json['macros']):
                if macro['name'] == new_macro['name']:  
                    current_json['macros'][idx] = new_macro
                    macro_exists = True
                    break
                
            if not macro_exists:
                current_json['macros'].append(new_macro) 
        else:
            current_json['macros'] = [new_macro]
    
        with open('macros.json', 'w') as file:
            json.dump(current_json, file, indent=4)
        
        MacroApp.refresh_buttons(self.mainWindow)
        self.on_close()
        
    def add_action(self, time, action_one=None, action_two=None):
        time = round(float(time), 2)
        if (time is not None) and not action_one and action_two is None:
            action = {
                "action": "sleep",
                "time": time
            }
            item = CustomFrameItem(self.scrollable_frame, label_text=f"Sleep | {time}", editor_instance=self, action_data=action)
        elif (time is not None) and action_one and action_two is None:
            action = {
                "action": "hold",
                "key": action_one,
                "time": time
            }
            item = CustomFrameItem(self.scrollable_frame, label_text=f"Hold | {time}", editor_instance=self, action_data=action, characters=f"{action_one}")
        else:
            action = {
                "action": "hold",
                "key": action_one,
                "keytwo": action_two,
                "time": time
            }
            item = CustomFrameItem(self.scrollable_frame, label_text=f"Hold | {time}", editor_instance=self, action_data=action, characters=f"{action_one}+{action_two}")

        item.pack(fill='x', padx=10, pady=5, expand=False)
        self.items.append(action)
        self.check_entries()
     
    
    def populate_existing_macro(self, existing_macro):
        self.nameInput.insert(0, existing_macro['name'])
        self.repetitionsInput.insert(0, str(existing_macro['repeatFor']))
        
        for action in existing_macro['actions']:
            if action['action'] == 'hold':
                if "keytwo" in action:
                    self.add_action(action['time'], action['key'], action['keytwo'])
                else:
                    self.add_action(action['time'], action['key'])
            elif action['action'] == 'sleep':
                self.add_action(action['time'])
    
    def validate_input(self, input_value):
        import re
        if input_value == "" or re.match(r'^-?\d*\.?\d*$', input_value):
            return True
        return False
    
    def on_close(self):
        self.destroy()
     
class MacroApp:
    
    def __init__(self):
        import threading

        self.toggle = False
        self.emergencyStop = False
        self.listener = None
        self.stop_event = threading.Event()
        self.selected_button_name = None
        self.app = ctk.CTk()
        self.startButton = None
        self.stopButton = None

        self.setup_ui()        
        self.start_listener_thread()

    def setup_ui(self):
        self.app.title("VS Macro app")
        self.app.geometry("600x500")
        TitleBar(self.app, on_close=self.on_close, title="VS Macro app", minimize=True).pack(fill="both")
    
        self.update_taskbar_icon()
        def resource_path(relative_path):
                   import os
                   import sys
                   try:
                       base_path = sys._MEIPASS
                   except Exception:
                       base_path = os.path.abspath(".")
                   return os.path.join(base_path, relative_path)
        self.app.iconbitmap(resource_path('logo.ico'))
        self.center_window()
    
        frame1 = ctk.CTkFrame(self.app, corner_radius=5, fg_color=DARK_GRAY)
        frame1.pack(expand=True, fill='both', padx=10, pady=10)
    
        left_frame = ctk.CTkFrame(frame1, width=300, corner_radius=0)
        left_frame.pack(side="left", fill="y", expand=False)
    
        self.scrollable_frame = ctk.CTkScrollableFrame(left_frame, corner_radius=4, label_text="Your macros:", label_fg_color=VERY_DARK_GRAY )
        self.scrollable_frame.pack(side="top", fill="both", expand=True)
    
        right_frame = ctk.CTkFrame(frame1, corner_radius=0)
        right_frame.pack(side="right", fill="y", expand=False)
    
        self.keybinds_file = load_data().get("keyBinds", {})
        self.startButton = self.keybinds_file.get("Start", "F8")
        self.stopButton = self.keybinds_file.get("Stop", "F9")
    
        self.start_button = ctk.CTkButton(right_frame, text=f"Start | {self.startButton}", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.start_macro)
        self.stop_button = ctk.CTkButton(right_frame, text=f"Stop | {self.stopButton}", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.emergency_stop)
        self.edit_button = ctk.CTkButton(right_frame, text="Edit", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.edit_macro)
        self.delete_button = ctk.CTkButton(right_frame, text="Delete", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.delete_macro)
        self.settings_button = ctk.CTkButton(right_frame, text="Settings", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , command=self.open_settings_popup)
    
        self.settings_button.pack(pady=5, padx=75, side="bottom")
        
        self.selected_button = [None]
        self.create_buttons()
        button_frame = ctk.CTkFrame(left_frame, corner_radius=0, fg_color=DARK_GRAY)
        button_frame.pack(side="top", fill="x")
    
        plus_button = ctk.CTkButton(button_frame, text="+", hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , width=30, command=self.open_editor_popup)
        plus_button.pack(side="top", padx=10, pady=5)
    
        self.app.bind_all(f"<{self.startButton}>", lambda event: self.start_macro())
        self.app.bind_all(f"<{self.stopButton}>", lambda event: self.emergency_stop()) 
    

    def update_taskbar_icon(self):
        window_id = win32gui.GetForegroundWindow()
        self.set_taskbar_icon(window_id)

    def center_window(self):
        screen_width = self.app.winfo_screenwidth()
        screen_height = self.app.winfo_screenheight()
        window_width, window_height = 600, 500
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.app.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

    def start_listener_thread(self):
        import threading

        self.listener_thread = threading.Thread(target=self.start_listener, daemon=True)
        self.listener_thread.start()

    def start_macro(self):
        import threading

        selected_macro = self.find_object_by_name(self.selected_button_name, load_data()["macros"])
        if selected_macro:
            self.toggle = True
            self.emergencyStop = False
            threading.Thread(target=lambda: self.execute_macro(selected_macro, selected_macro["repeatFor"]), daemon=True).start()

    def execute_macro(self, macro, macroRepeats):
        import time
        timesDone = 0
        while self.toggle and not self.emergencyStop:
            if macroRepeats == -1 or timesDone < macroRepeats:
                timesDone += 1
                self.execute_actions(macro)
            else:
                self.toggle = False
            time.sleep(0.01)

    def emergency_stop(self):
        self.emergencyStop = True
        self.toggle = False

    def execute_actions(self, macro):
        import time
        for action in macro['actions']:
            if self.emergencyStop:
                break
            if action['action'] == 'hold':
                try:
                    self.macroCommand(action['key'], action['time'], action.get('keytwo'))
                except Exception:
                    self.macroCommand(action['key'], action['time'])
            elif action['action'] == 'sleep':
                time.sleep(action['time'])

    def macroCommand(self, key: str, holdTime: float, holdClick: str = None):
        import time
        keyDown(key)
        if holdClick:
            keyDown(holdClick)
        start_time = time.time()
        
        while time.time() - start_time < holdTime and not self.emergencyStop:
            time.sleep(0.001)
        
        keyUp(key)
        if holdClick:
            keyUp(holdClick)

    def emergency_stop(self):
        self.emergencyStop = True
        self.toggle = False

    def start_listener(self):
        self.stop_event.clear()
        with keyboard.Listener(on_press=lambda key: self.on_press(key) if not self.stop_event.is_set() else None) as self.listener:
            self.listener.join()

    def stop_listener(self):
        if self.listener:
            self.stop_event.set()
            self.listener.stop()
            self.listener = None

    def restart_listener(self):
        import threading

        self.stop_listener()
        threading.Thread(target=self.start_listener, daemon=True).start()

    def on_press(self, key):
        try:
            if key == self.strToKey(self.startButton):
                self.toggle = not self.toggle
                if self.toggle:
                    self.start_macro()
            elif key == self.strToKey(self.stopButton):
                self.emergencyStop = True
                self.toggle = False
        except AttributeError:
            pass

    def create_buttons(self):
        for macro in load_data()["macros"]:
            button = ctk.CTkButton(self.scrollable_frame, text=macro["name"], hover_color=ALMOST_BLACK , fg_color=VERY_DARK_GRAY , font=("font", 14))
            button.configure(command=lambda b=button: self.select_macro(b))
            button.pack(pady=5, padx=5)

    def select_macro(self, button):
        if self.selected_button[0] == button:
            button.configure(fg_color=VERY_DARK_GRAY )
            button.configure(font=("font", 14))
            self.start_button.pack_forget()
            self.stop_button.pack_forget()
            self.edit_button.pack_forget()
            self.delete_button.pack_forget()
            self.selected_button[0] = None
        else:
            if self.selected_button[0]:
                self.selected_button[0].configure(fg_color=VERY_DARK_GRAY )
                self.selected_button[0].configure(font=("font", 14))
            button.configure(fg_color=ALMOST_BLACK )
            button.configure(font=("font", 16))
            self.selected_button[0] = button
            self.start_button.pack(pady=(45, 5), padx=75, side="top")
            self.stop_button.pack(pady=0, padx=75, side="top")
            self.edit_button.pack(pady=(35, 5), padx=75, side="top")
            self.delete_button.pack(pady=0, padx=75, side="top")
            self.selected_button_name = button.cget("text")

    def reset_selection_on_focus(self):
        if self.selected_button[0]:
            try:
                self.selected_button[0].configure(fg_color=VERY_DARK_GRAY )
                self.selected_button[0].configure(font=("font", 14))
            except:
                pass
            self.selected_button[0] = None

        self.start_button.pack_forget()
        self.stop_button.pack_forget()
        self.edit_button.pack_forget()
        self.delete_button.pack_forget()

    def delete_macro(self):
        current_data = load_data()
        current_data['macros'] = [macro for macro in current_data['macros'] if macro['name'] != self.selected_button_name]
        with open("macros.json", 'w') as file:
            json.dump(current_data, file, indent=4)
        self.clear_frame(self.scrollable_frame)
        self.create_buttons()
        self.start_button.pack_forget()
        self.stop_button.pack_forget()
        self.edit_button.pack_forget()
        self.delete_button.pack_forget()

    def clear_frame(self, frame):
        for widget in frame.winfo_children():
            widget.pack_forget()

    def open_settings_popup(self):
        self.app.withdraw() 
        popup = SettingsPopup(self.app, self,currentstart=self.startButton,currentstop=self.stopButton)
        self.app.wait_window(popup) 
        self.app.deiconify() 

    def open_editor_popup(self):
        self.app.withdraw() 
        popup = EditorPopup(self.app,self)
        self.app.wait_window(popup) 
        self.app.deiconify()    

    def edit_macro(self):
        existing_macro = self.find_object_by_name(self.selected_button_name, load_data()["macros"])
        self.app.withdraw() 
        popup = EditorPopup(self.app,self, existing_macro=existing_macro)
        self.app.wait_window(popup) 
        self.app.deiconify() 
        
    def update_button_text(self):
        self.start_button.configure(text=f"Start | {self.startButton}")
        self.stop_button.configure(text=f"Stop | {self.stopButton}")
        
    def update_keybind(self,new_start_key, new_stop_key):
        self.startButton = new_start_key
        self.stopButton = new_stop_key  
        self.data = load_data()
        if new_start_key:
            self.data['keyBinds']['Start'] = new_start_key
        if new_stop_key:
            self.data['keyBinds']['Stop'] = new_stop_key
        with open("macros.json", 'w') as file:
            json.dump(self.data, file, indent=4)
    def refresh_buttons(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.reset_selection_on_focus() 
        self.create_buttons()
        
    def strToKey(self,s):
        return getattr(keyboard.Key, s.lower()) if len(s) == 2 else keyboard.KeyCode.from_char(s)

    def find_object_by_name(self,name, macros):
        return next((macro for macro in macros if macro.get('name') == name), None)
    
    def set_taskbar_icon(self,window_id):
        import win32api
        current_style = win32api.GetWindowLong(window_id, -20)
        new_style = (current_style | 0x00040000) & ~0x00000080
        win32api.SetWindowLong(window_id, -20, new_style)
        win32gui.SetWindowPos(window_id, win32con.HWND_TOP, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER)
        
    def on_close(self):
        self.app.destroy()
def main():
    app = MacroApp()
    app.app.mainloop()

if __name__ == "__main__":
    main()
