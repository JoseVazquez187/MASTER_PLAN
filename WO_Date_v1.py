import json
import os
import time
import pyautogui
import win32gui
import win32con
import win32api
import pandas as pd

cpu = f"{os.environ['USERNAME']}"
MAIN_WINDOW_TITLE_PARTIAL = "Royal 4 - Live - K-95"
USERNAME_WINDOW_TITLE_PARTIAL = f"Name ({cpu})"
PASSWORD_WINDOW_TITLE_PARTIAL = "Password"
config_file_path = "C:\\Temp\\Smart_Tools\\myenv\\config.json" 
excel_file_path = "C:\\Users\\J.Vazquez\\Desktop\\WO to change.xlsx"

# Load the JSON data
with open(config_file_path, 'r') as file:
    config_data = json.load(file)

username = config_data.get('username')
password = config_data.get('password')

# Function to ensure Caps Lock is off
def ensure_capslock_off():
    if win32api.GetKeyState(0x14) == 1:  # 0x14 is the virtual key code for Caps Lock
        win32api.keybd_event(0x14, 0, 0, 0)  # Turn off Caps Lock
        time.sleep(0.1)  # Small delay to ensure Caps Lock turns off

# Function to print all visible windows' titles for debugging
def list_open_windows():
    def enum_window_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            if window_title:  # Only print windows with a title
                results.append(window_title)

    windows = []
    win32gui.EnumWindows(enum_window_callback, windows)
    print("Open Windows:")
    for title in windows:
        print(f"  - {title}")

# Function to find a window by partial title and bring it to the front
def find_window_by_partial_title(partial_title, max_wait=10):
    hwnd = None
    waited_time = 0
    
    def enum_window_callback(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            if partial_title.lower() in window_title.lower():
                results.append(hwnd)

    while not hwnd and waited_time < max_wait:
        results = []
        win32gui.EnumWindows(enum_window_callback, results)
        
        if results:
            hwnd = results[0]
            win32gui.SetForegroundWindow(hwnd)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            return hwnd
        else:
            time.sleep(1)
            waited_time += 1

    print(f"Window containing '{partial_title}' not found!")
    return None

# Function to check for continue dialog (simple text detection)
def check_for_continue_dialog():
    # Usar pyautogui para detectar texto en pantalla o simplemente 
    # asumir que después de cierto tiempo puede aparecer
    time.sleep(0.5)
    return False  # Simplificado por ahora

# Function to handle continue dialog
def handle_continue_dialog():
    print("Handling 'Do you want to continue?' dialog...")
    pyautogui.typewrite('y')
    pyautogui.press('enter')
    time.sleep(1)

def start_r4():
    # List all open windows for debugging
    list_open_windows()

    # Path to the application you want to open
    kermit_path = r"C:\Program Files (x86)\Kermit 95 2.1\royal4.ksc"
    os.popen('"{}"'.format(kermit_path))

    # Wait and bring the main application window to the front
    main_hwnd = find_window_by_partial_title(MAIN_WINDOW_TITLE_PARTIAL)
    if not main_hwnd:
        return

    # Wait and bring the username window to the front
    username_hwnd = find_window_by_partial_title(USERNAME_WINDOW_TITLE_PARTIAL)
    if not username_hwnd:
        return

    # Ensure Caps Lock is off and type the username
    time.sleep(1)
    ensure_capslock_off()
    pyautogui.typewrite(username)
    pyautogui.press('enter')  # Press Enter to submit the username

    # Wait and bring the password window to the front
    password_hwnd = find_window_by_partial_title(PASSWORD_WINDOW_TITLE_PARTIAL)
    if not password_hwnd:
        return

    # Ensure Caps Lock is off and type the password
    ensure_capslock_off()
    pyautogui.typewrite(password)
    pyautogui.press('enter')  # Press Enter to submit the password

    # Additional navigation after login
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(.5)
    pyautogui.typewrite('pr 4 1')
    time.sleep(.5)
    pyautogui.press('enter')
    time.sleep(.5)
    pyautogui.press('enter')
    time.sleep(.5)


    pyautogui.press('enter')
    time.sleep(.5)
    pyautogui.press('enter')

    # CARGAR EL DATAFRAME Y PROCESAR CADA WO
    print("Loading Excel file...")
    try:
        df = pd.read_excel(excel_file_path)
        print(f"Loaded {len(df)} rows from Excel file")
        print("Columns in file:", df.columns.tolist())
        
        # Iterar sobre cada fila del dataframe
        for index, row in df.iterrows():
            wo_number = row['WO No']  # Asumiendo que la columna se llama 'WO No'
            print(f"\nProcessing WO {index + 1}/{len(df)}: {wo_number}")
            
            # Limpiar campo y escribir el número de WO
            ensure_capslock_off()
            pyautogui.typewrite(str(wo_number))
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(2)  # Esperar a que cargue la WO
            
            # Imprimir la WO

            
            # Dar 7 enters y manejar posibles diálogos
            print("Pressing 6 enters...")
            for i in range(6):
                pyautogui.press('enter')
                time.sleep(0.5)

            # Esperar y responder a posible diálogo
            print("Checking for dialog...")
            time.sleep(1)
            pyautogui.typewrite('y')
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(0.5)

            
            # Manejar segundo diálogo posible (fecha past due)
            print("Checking for past due date dialog...")
            time.sleep(1)
            pyautogui.typewrite('y')
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(0.5)


            print(f"Completed WO {wo_number}")
            time.sleep(1)
            
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {excel_file_path}")
        return
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        return
    
    print("\nAll work orders have been processed!")

if __name__ == "__main__":
    start_r4()