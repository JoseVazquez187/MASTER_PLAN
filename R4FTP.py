import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os
from pathlib import Path
import time

def validar_campos(*args):
    if entry_user.get() and entry_pass.get():
        btn_download.config(state="normal")
    else:
        btn_download.config(state="disabled")

def conectar_y_descargar():
    user = entry_user.get()
    passwd = entry_pass.get()
    server = combo_server.get()

    escritorio = Path.home() / "Desktop"
    script_path = escritorio / "ftp_script.txt"

    # Crear script FTP
    with open(script_path, "w") as f:
        f.write(f"open {server}\n")
        f.write(f"user {user} {passwd}\n")
        f.write(f"get {user}.txt\n")
        f.write("bye\n")

    os.chdir(escritorio)

    # Mostrar mensaje de espera
    status_label.config(text="Conectando y descargando archivo...")
    root.update()  # Actualiza la interfaz antes de ejecutar

    # Medir tiempo
    start = time.time()
    subprocess.call(f'start /wait cmd /c ftp -n -s:"{script_path.name}"', shell=True)
    end = time.time()

    duracion = round(end - start, 2)
    status_label.config(text=f"Transferencia completada en {duracion} segundos.")

# Crear ventana
root = tk.Tk()
root.title("EZAirR4FTP")
root.geometry("310x330")
root.resizable(False, False)

# Royal4 Server
tk.Label(root, text="Royal4 Server:").place(x=20, y=20)
combo_server = ttk.Combobox(root, values=["11.62.16.159", "Test", "Dev"], state="readonly", width=20)
combo_server.place(x=130, y=20)
combo_server.current(0)

# Usuario
tk.Label(root, text="Royal4 User Name:").place(x=20, y=60)
entry_user = tk.Entry(root, width=25)
entry_user.place(x=130, y=60)

# Contraseña
tk.Label(root, text="Royal4 Password:").place(x=20, y=100)
entry_pass = tk.Entry(root, show="*", width=25)
entry_pass.place(x=130, y=100)

# Tipo de transferencia
tk.Label(root, text="Transfer Type:").place(x=20, y=140)
transfer_type = tk.StringVar(value="ASCII")
tk.Radiobutton(root, text="ASCII", variable=transfer_type, value="ASCII").place(x=130, y=140)
tk.Radiobutton(root, text="Binary", variable=transfer_type, value="Binary").place(x=190, y=140)

# Advertencia
label_warning = tk.Label(
    root,
    text="Please be sure that the password you type above\n"
         "is correct. Three tries with the wrong password\n"
         "will cause you to be locked out of Royal4 until\n"
         "you contact the help desk!",
    fg="red", justify="left", wraplength=280
)
label_warning.place(x=15, y=180)

# Botón de descarga
btn_download = tk.Button(root, text="Download", width=20, command=conectar_y_descargar, state="disabled")
btn_download.place(x=80, y=260)

# Mensaje de estado
status_label = tk.Label(root, text="", fg="blue")
status_label.place(x=40, y=295)

# Help (opcional)
btn_help = tk.Button(root, text="Help")
btn_help.place(x=260, y=290)

# Activar validación de campos
entry_user.bind("<KeyRelease>", validar_campos)
entry_pass.bind("<KeyRelease>", validar_campos)

root.mainloop()
