import tkinter as tk
from tkinter import ttk
import pyautogui
import threading
import time

class ScreenKeepAlive:
    def __init__(self, root):
        self.root = root
        self.root.title("Keep Screen Awake")
        self.root.geometry("500x400")
        self.root.minsize(450, 350)
        self.root.resizable(True, True)
        
        self.is_running = False
        self.thread = None
        
        self.setup_ui()
        
        # Configurar el grid para que se expanda
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar el grid del frame principal
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="🖥️ Keep Screen Awake", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky=tk.W+tk.E)
        
        # Configuración en un frame separado
        config_frame = ttk.LabelFrame(main_frame, text="Configuración", padding="10")
        config_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        config_frame.columnconfigure(1, weight=1)
        
        # Intervalo de tiempo
        ttk.Label(config_frame, text="Intervalo (segundos):").grid(row=0, column=0, sticky=tk.W, pady=8)
        self.interval_var = tk.StringVar(value="60")
        interval_entry = ttk.Entry(config_frame, textvariable=self.interval_var, width=15)
        interval_entry.grid(row=0, column=1, sticky=tk.W, padx=(15, 0), pady=8)
        
        # Método de prevención
        ttk.Label(config_frame, text="Método:").grid(row=1, column=0, sticky=tk.W, pady=8)
        self.method_var = tk.StringVar(value="mouse")
        method_combo = ttk.Combobox(config_frame, textvariable=self.method_var,
                                   values=["mouse", "key"], state="readonly", width=20)
        method_combo.grid(row=1, column=1, sticky=tk.W, padx=(15, 0), pady=8)
        
        # Botones
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=15)
        
        self.start_button = ttk.Button(button_frame, text="▶️ Iniciar", 
                                      command=self.start_keepalive, width=12)
        self.start_button.pack(side=tk.LEFT, padx=(0, 15))
        
        self.stop_button = ttk.Button(button_frame, text="⏹️ Detener", 
                                     command=self.stop_keepalive, state=tk.DISABLED, width=12)
        self.stop_button.pack(side=tk.LEFT)
        
        # Estado
        status_frame = ttk.LabelFrame(main_frame, text="Estado", padding="10")
        status_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.status_var = tk.StringVar(value="⏸️ Detenido")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                font=("Arial", 11, "bold"))
        status_label.pack()
        
        # Log
        log_frame = ttk.LabelFrame(main_frame, text="Registro de actividad", padding="10")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 5))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Text widget con scrollbar
        text_frame = ttk.Frame(log_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(text_frame, height=10, width=50, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
    def log_message(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def start_keepalive(self):
        if self.is_running:
            return
            
        try:
            interval = int(self.interval_var.get())
            if interval <= 0:
                raise ValueError("El intervalo debe ser mayor a 0")
        except ValueError as e:
            self.log_message(f"❌ Error: {e}")
            return
            
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_var.set("▶️ Ejecutándose")
        
        self.log_message(f"🚀 Iniciado - Método: {self.method_var.get()}, Intervalo: {interval}s")
        
        self.thread = threading.Thread(target=self.keepalive_worker, args=(interval,), daemon=True)
        self.thread.start()
        
    def stop_keepalive(self):
        if not self.is_running:
            return
            
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_var.set("⏸️ Detenido")
        
        self.log_message("🛑 Detenido por el usuario")
        
    def keepalive_worker(self, interval):
        while self.is_running:
            try:
                if self.method_var.get() == "mouse":
                    # Mover mouse
                    current_x, current_y = pyautogui.position()
                    pyautogui.move(1, 1)
                    time.sleep(0.1)
                    pyautogui.move(-1, -1)
                    self.log_message(f"🖱️ Mouse movido en ({current_x}, {current_y})")
                else:
                    # Presionar tecla F15 (no hace nada visible)
                    pyautogui.press('f15')
                    self.log_message("⌨️ Tecla F15 presionada")
                    
                time.sleep(interval)
                
            except Exception as e:
                self.log_message(f"❌ Error: {e}")
                self.is_running = False
                self.root.after(0, self.stop_keepalive)
                break

def main():
    root = tk.Tk()
    app = ScreenKeepAlive(root)
    
    # Manejar cierre de ventana
    def on_closing():
        app.stop_keepalive()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()