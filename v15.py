import os
import keyboard
import time
import subprocess
import tkinter as tk    
import threading
import json     
import pyautogui
import sys

# Initialize stop event for key listener control
stop_event = threading.Event()      

# Load paths from the JSON configuration file
config_path = "C:\\Temp\\Smart_Tools\\myenv\\config.json"
with open(config_path, "r") as config_file:
    config = json.load(config_file)

interpreter_path = config.get("interpreter_path")
root_folder_path = config.get("folder_path")  # Base path: C:\Temp\Smart_Tools
current_path = root_folder_path  # Start at the root folder

# Set constants for page and button display limits
buttons_per_page = 5
page_start = 0                       

control_window = None
current_display = "folders"

# Function to close all Tkinter windows and exit the application
def close_all_windows():
    global control_window
    stop_event.set()  # Signal to stop any running threads

    if control_window and control_window.winfo_exists():
        control_window.destroy()

    sys.exit()

# Confirmation dialog function
def show_confirmation(root, x_position, y_position):
    confirmation_window = tk.Toplevel(root)
    confirmation_window.overrideredirect(True)
    confirmation_window.attributes("-topmost", True)
    confirmation_window.wm_attributes("-alpha", 0.8)  # Set transparency to 80%
    confirmation_window.focus_force()

    confirmation_window_width, confirmation_window_height = 200, 100
    confirmation_y_position = y_position - confirmation_window_height - 10
    confirmation_window.geometry(f"{confirmation_window_width}x{confirmation_window_height}+{x_position}+{confirmation_y_position}")

    frame = tk.Frame(confirmation_window, bg="#333333", bd=5, relief="raised")
    frame.pack(fill="both", expand=True)

    label = tk.Label(frame, text="Are you sure you want to close?", font=("Arial", 10, "bold"), bg="#333333", fg="white", wraplength=180)
    label.pack(pady=(10, 5))

    button_frame = tk.Frame(frame, bg="#333333")
    button_frame.pack(pady=5)

    def on_focus(event):
        event.widget.config(bg="#3399FF", fg="white")

    def on_focus_out(event):
        event.widget.config(bg="#444444", fg="white")

    yes_button = tk.Button(button_frame, text="Yes", width=8, command=close_all_windows, bg="#444444", activebackground="#4CAF50", fg="white")
    yes_button.grid(row=0, column=0, padx=5)
    yes_button.bind("<FocusIn>", on_focus)
    yes_button.bind("<FocusOut>", on_focus_out)

    no_button = tk.Button(button_frame, text="No", width=8, command=confirmation_window.destroy, bg="#444444", activebackground="#3399FF", fg="white")
    no_button.grid(row=0, column=1, padx=5)
    no_button.bind("<FocusIn>", on_focus)
    no_button.bind("<FocusOut>", on_focus_out)

    yes_button.focus_set()
    confirmation_window.bind("<Escape>", lambda e: confirmation_window.destroy())

# Function to create the small Tkinter window showing "Running" status
def create_running_window():
    root = tk.Tk()
    root.title("Status")

    window_width, window_height = 200, 50
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x_position = screen_width - window_width - 45
    y_position = screen_height - window_height - 60
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    root.attributes("-topmost", True)
    root.overrideredirect(True)

    main_frame = tk.Frame(root, bg="#d90467", bd=5, relief="raised")
    main_frame.place(relx=0.5, rely=0.5, anchor="center", width=window_width, height=window_height)

    status_label = tk.Label(main_frame, text="Running", font=("Arial", 12, "bold"), bg="#d90467", fg="white")
    status_label.pack(expand=True, fill=tk.BOTH)

    root.bind("<Escape>", lambda e: show_confirmation(root, x_position, y_position))

    root.mainloop()

# Function to highlight button on hover
def on_hover(event):
    event.widget.config(bg="#4CAF50")  # Green highlight color

# Function to reset button color on hover exit
def off_hover(event):
    event.widget.config(bg="#666666")  # Original button color

# Function to highlight button on focus
def on_focus(event):
    event.widget.config(bg="#4CAF50")  # Green color on focus

# Function to reset button color on focus out
def on_focus_out(event):
    event.widget.config(bg="#666666")  # Original button color

# Function to run the control window
def run_main():
    global control_window, current_display, folder_buttons, file_buttons, page_start, nav_frame, current_path

    folder_buttons = []
    file_buttons = []
    page_start = 0

    if control_window is not None and control_window.winfo_exists():
        control_window.lift()
        return

    control_window = tk.Tk()
    control_window.overrideredirect(True)
    control_window.attributes("-topmost", True)
    control_window.configure(bg="black")
    control_window.wm_attributes("-transparentcolor", "black", "-alpha", 0.8)

    screen_width, screen_height = control_window.winfo_screenwidth(), control_window.winfo_screenheight()
    window_width, window_height = 200, 300
    x_position = screen_width - window_width - 50
    y_position = screen_height - window_height - 150
    control_window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    # Separate nav_frame for Back, Previous, and Next buttons to ensure visibility
    nav_frame = tk.Frame(control_window, bg="black", width=window_width)
    prev_button = tk.Button(nav_frame, text="Previous", bg="#666666", fg="white", width=8, command=previous_page)
    next_button = tk.Button(nav_frame, text="Next", bg="#666666", fg="white", width=8, command=next_page)
    back_button = tk.Button(nav_frame, text="Back", bg="#666666", fg="white", width=8, command=navigate_back)

    # Bind hover and focus events to the buttons for highlight effect
    for button in [prev_button, next_button, back_button]:
        button.bind("<Enter>", on_hover)
        button.bind("<Leave>", off_hover)
        button.bind("<FocusIn>", on_focus)
        button.bind("<FocusOut>", on_focus_out)

    # Arrange buttons vertically to ensure they are visible
    prev_button.pack(side="top", fill="x", pady=2)
    next_button.pack(side="top", fill="x", pady=2)
    back_button.pack(side="top", fill="x", pady=2)

    nav_frame.pack(side="right", fill="y", padx=5, pady=5)

    def close_window(event=None):
        global control_window
        if control_window:
            control_window.destroy()
            control_window = None

    control_window.bind("<Escape>", close_window)
    control_window.protocol("WM_DELETE_WINDOW", close_window)

    display_items(current_path)

    # Automatically focus and simulate a click on the first button
    if folder_buttons:
        folder_buttons[0].focus_set()
        control_window.update()
        x, y = folder_buttons[0].winfo_rootx() + 5, folder_buttons[0].winfo_rooty() + 5
        pyautogui.click(x, y)  # Simulate a click on the first button

    control_window.mainloop()

def create_button(text, command):
    # Create the button with the specified style and assign the command directly to it
    btn = tk.Button(
        control_window, 
        text=text, 
        bg="#444444", 
        activebackground="#4CAF50", 
        fg="white",
        font=("Arial", 10, "bold"), 
        height=2, 
        width=20, 
        command=command  # Directly set the command for the button click
    )
    
    # Bind "<Button-1>" to disable mouse clicks (if this is desired behavior)
    btn.bind("<Button-1>", disable_click)  # Disables mouse click on the button
    btn.bind("<FocusIn>", on_focus)  # Highlights the button when it gains focus
    btn.bind("<FocusOut>", on_focus_out)  # Resets the button's appearance when it loses focus
    
    # Setting focus on the button so that it listens for keyboard input, such as the Space key
    btn.focus_set()  # Ensure the button is initially focused

    # Return the button object to be added to the Tkinter window
    return btn

# Function to disable mouse clicks
def disable_click(event):
    return "break"

# Function to display both folders and files in the current directory
def display_items(path):
    global current_display, folder_buttons, file_buttons, page_start, current_path
    current_display = "items"
    current_path = path
    page_start = 0  # Reset page_start when displaying a new set of items

    # Clear previous buttons
    for widget in control_window.winfo_children():
        if isinstance(widget, tk.Button) and widget not in nav_frame.winfo_children():
            widget.pack_forget()

    # Get items in the path and recreate buttons
    items = [f for f in os.listdir(path)]
    folder_buttons = [create_button(item, lambda fn=os.path.join(path, item): open_item(fn)) for item in items]

    # folder_buttons = [create_button(item, lambda e, fn=os.path.join(path, item): open_item(fn)) for item in items]
    update_buttons(folder_buttons)  # Display buttons for the current page

# Function to go back one level in the path or to the previous page if not on the first page
def navigate_back():
    global current_path, page_start, folder_buttons
    if page_start > 0:
        # If not on the first page, go to the previous page
        page_start -= buttons_per_page
        update_buttons(folder_buttons)
    elif current_path != root_folder_path:
        # If on the first page and not in the root directory, go up one level
        current_path = os.path.dirname(current_path)
        page_start = 0  # Reset page_start to start from the beginning
        display_items(current_path)  # Refresh display with the updated path

# Helper function to open folder or file based on type
def open_item(path):
    if os.path.isdir(path):
        display_items(path)
    else:
        run_file(path)

# Function to highlight button on focus with color differentiation for folders and files
def on_focus(event):
    button_text = event.widget.cget("text")
    path = os.path.join(current_path, button_text)
    if os.path.isdir(path):
        event.widget.config(bg="orange", fg="white")
    else:
        event.widget.config(bg="#1E90FF", fg="white")

# Function to reset button color on focus out
def on_focus_out(event):
    event.widget.config(bg="#444444", fg="white")

# Function to update displayed buttons
def update_buttons(button_list):
    # Clear any existing displayed buttons (except navigation)
    for widget in control_window.winfo_children():
        if isinstance(widget, tk.Button) and widget not in nav_frame.winfo_children():
            widget.pack_forget()

    # Display the correct range of buttons for the current page
    for btn in button_list[page_start:page_start + buttons_per_page]:
        btn.pack(fill=tk.X, pady=2, padx=5)

# Pagination functions
def next_page():
    global page_start
    if page_start + buttons_per_page < len(folder_buttons):
        page_start += buttons_per_page
        update_buttons(folder_buttons)

def previous_page():
    global page_start
    if page_start - buttons_per_page >= 0:
        page_start -= buttons_per_page
        update_buttons(folder_buttons)

# Run file if it's a .py or .exe
def run_file(file_path):
    if file_path.endswith(".py"):
        subprocess.run([interpreter_path, file_path])
    #elif file_path.endswith(".exe"):
        #subprocess.run(file_path)

# Key listener function
def key_listener():

    try:
        while not stop_event.is_set():
            if keyboard.is_pressed("ctrl") and keyboard.is_pressed("alt") and keyboard.is_pressed("1"):

                run_main()
                time.sleep(0.5)
            time.sleep(0.1)
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    threading.Thread(target=create_running_window, daemon=True).start()
    key_listener()
