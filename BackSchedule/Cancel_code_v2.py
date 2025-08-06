import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, BooleanVar, Checkbutton
from tkinter.scrolledtext import ScrolledText
import threading
import sqlite3
from tqdm import tqdm

# === Interfaz Tkinter ===
root = tk.Tk()
root.title("Carga y Validación de Cancelaciones")
root.geometry("900x600")

frame_checkboxes = tk.Frame(root)
frame_checkboxes.pack(pady=5)

tk.Label(frame_checkboxes, text="Selecciona los códigos de acción que deseas rastrear:", font=("Segoe UI", 10, "bold")).pack()

codigos_posibles = ["RI", "PD", "PR", "RO", "CN", "AD"]
codigos_dict = {}

for codigo in codigos_posibles:
    var = BooleanVar(value=(codigo == "CN"))
    chk = Checkbutton(frame_checkboxes, text=codigo, variable=var, font=("Segoe UI", 10))
    chk.pack(side="left", padx=10)
    codigos_dict[codigo] = var

log_output = ScrolledText(root, font=("Consolas", 10))
log_output.pack(expand=True, fill="both")

# === Función para insertar texto en la interfaz ===
def log_insert(mensaje):
    log_output.insert(tk.END, mensaje + "\n")
    log_output.see(tk.END)
    root.update_idletasks()

# === Función para buscar la fila del encabezado ===
def cargar_con_encabezado_dinamico(path_archivo):
    df_raw = pd.read_excel(path_archivo, header=None, usecols="A:K")
    coincidencias = df_raw.apply(lambda row: row.astype(str).str.contains("ITEM-NO", case=False)).any(axis=1)
    if not coincidencias.any():
        raise ValueError("No se encontró la fila con encabezado 'ITEM-NO'")
    fila_header = coincidencias[coincidencias].index[0]
    df = pd.read_excel(path_archivo, header=fila_header)
    return df

# === Procesar archivo e insertar ===
def procesar_archivo_y_insertar(ruta_archivo, conn, archivo, codigos_seleccionados, timestamp, archivos_previos):
    archivo_normalizado = archivo.strip().upper()
    if archivo_normalizado in archivos_previos:
        return f"⏭️ {archivo} ya fue procesado previamente. Se omite."

    try:
        fecha_archivo = datetime.strptime(archivo.replace("Action-", "").replace(".xlsx", ""), "%Y%m%d")
        df = cargar_con_encabezado_dinamico(ruta_archivo)

        if 'REF' not in df.columns or 'SUB' not in df.columns:
            return f"⚠️ {archivo} omitido (sin columnas 'REF' o 'SUB')."

        a_cd_column = next((col for col in df.columns if str(col).strip().upper() == "A-CD"), None)
        if not a_cd_column:
            return f"⚠️ {archivo} omitido (sin columna 'A-CD')."

        # Limpiar comillas simples de todas las columnas de texto
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip().str.replace("'", "")
        
        df['REF'] = df['REF'].astype(str).str.strip().str.replace("'", "")
        df['SUB'] = df['SUB'].astype(str).str.strip().str.replace("'", "")
        df['PO_Linea'] = df['REF'] + '-' + df['SUB']
        df['A-CD'] = df[a_cd_column].astype(str).str.upper().str.strip().str.replace("'", "")
        df['Fecha_Archivo'] = fecha_archivo.strftime("%Y-%m-%d")

        df_filtrado = df[df['A-CD'].isin(codigos_seleccionados)][['PO_Linea', 'A-CD', 'Fecha_Archivo']]
        if not df_filtrado.empty:
            # Limpiar comillas simples de todas las columnas en df_filtrado
            for col in df_filtrado.columns:
                if df_filtrado[col].dtype == 'object':
                    df_filtrado[col] = df_filtrado[col].astype(str).str.replace("'", "")
            df_filtrado.to_sql("first_action_cancel", conn, if_exists="append", index=False)

        cursor = conn.cursor()
        cursor.execute("INSERT INTO files_action_checked (archivo, fecha) VALUES (?, ?)", (archivo_normalizado, timestamp))
        conn.commit()

        return f"✅ {archivo} insertado con {len(df_filtrado)} registros."

    except Exception as e:
        return f"❌ Error procesando {archivo}: {e}"

# === Función principal ===
def ejecutar_proceso():
    codigos_seleccionados = [codigo for codigo, var in codigos_dict.items() if var.get()]
    if not codigos_seleccionados:
        messagebox.showwarning("Sin selección", "Selecciona al menos un código de acción.")
        return

    log_insert("🟡 Selecciona el archivo de cancelaciones base")
    archivo_cancelaciones = filedialog.askopenfilename(title="Selecciona el archivo de cancelaciones", filetypes=[("Excel files", "*.xlsx")])
    if not archivo_cancelaciones:
        log_insert("❌ No se seleccionó archivo base. Cancelando.")
        return

    # Solo usar R4Database
    r4_db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    log_insert("🟡 Selecciona la carpeta del histórico de actions")
    path_historico = filedialog.askdirectory(title="Selecciona carpeta del histórico de actions")
    if not path_historico:
        log_insert("❌ No se seleccionó carpeta de histórico. Cancelando.")
        return

    # Crear directorio si no existe
    r4_dir = os.path.dirname(r4_db_path)
    if not os.path.exists(r4_dir):
        os.makedirs(r4_dir)
        log_insert(f"📁 Directorio creado: {r4_dir}")

    # Conectar directamente a R4Database
    conn = sqlite3.connect(r4_db_path)
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS first_action_cancel (PO_Linea TEXT, 'A-CD' TEXT, Fecha_Archivo TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS files_action_checked (archivo TEXT PRIMARY KEY, fecha TEXT)")
    conn.commit()

    log_insert(f"📊 Conectado a R4Database: {r4_db_path}")

    archivos = sorted(os.listdir(path_historico))
    cursor.execute("SELECT archivo FROM files_action_checked")
    archivos_previos = set(row[0].strip().upper() for row in cursor.fetchall())

    total_insertados = 0
    for archivo in tqdm(archivos, desc="Insertando archivos", unit="archivo"):
        if archivo.startswith("Action-") and archivo.endswith(".xlsx"):
            ruta_archivo = os.path.join(path_historico, archivo)
            resultado = procesar_archivo_y_insertar(ruta_archivo, conn, archivo, codigos_seleccionados, timestamp, archivos_previos)
            log_insert(resultado)

    # Resultado Final
    try:
        df_cancelaciones = pd.read_excel(archivo_cancelaciones, header=None)
        coincidencias = df_cancelaciones.apply(lambda row: row.astype(str).str.contains("REF", case=False)).any(axis=1)
        fila_encabezado = coincidencias[coincidencias].index[0]
        df_cancelaciones = pd.read_excel(archivo_cancelaciones, header=fila_encabezado)

        # Limpiar comillas simples de todas las columnas de texto del archivo de cancelaciones
        for col in df_cancelaciones.columns:
            if df_cancelaciones[col].dtype == 'object':
                df_cancelaciones[col] = df_cancelaciones[col].astype(str).str.strip().str.replace("'", "")
        
        df_cancelaciones['REF'] = df_cancelaciones['REF'].astype(str).str.strip().str.replace("'", "")
        df_cancelaciones['SUB'] = df_cancelaciones['SUB'].astype(str).str.strip().str.replace("'", "")
        df_cancelaciones['PO_Linea'] = df_cancelaciones['REF'] + '-' + df_cancelaciones['SUB']
        df_cancelaciones['A-CD'] = df_cancelaciones['A-CD'].astype(str).str.upper().str.strip().str.replace("'", "")
        df_cancelaciones = df_cancelaciones[df_cancelaciones['A-CD'].isin(codigos_seleccionados)]

        query = """
            SELECT PO_Linea, "A-CD", MIN(Fecha_Archivo) AS Fecha_Primera_Aparicion
            FROM first_action_cancel
            GROUP BY PO_Linea, "A-CD"
        """
        df_sqlite = pd.read_sql_query(query, conn)

        df_resultado = df_cancelaciones.merge(df_sqlite, on=['PO_Linea', 'A-CD'], how='left')
        df_resultado['Fecha_Primera_Aparicion'] = pd.to_datetime(df_resultado['Fecha_Primera_Aparicion'], errors='coerce')
        df_resultado['Días_Sin_Ejecutar'] = (datetime.today() - df_resultado['Fecha_Primera_Aparicion']).dt.days
        
        # Limpiar comillas simples de todas las columnas antes del insert final
        for col in df_resultado.columns:
            if df_resultado[col].dtype == 'object':
                df_resultado[col] = df_resultado[col].astype(str).str.replace("'", "")
        
        df_resultado.to_sql("cancel_action_aging", conn, if_exists="replace", index=False)
        log_insert("📁 Resultado final insertado en R4Database.db")
            
    except Exception as e:
        log_insert(f"❌ Error generando resultado final: {e}")

    conn.commit()
    conn.close()
    log_insert("✅ Proceso completado exitosamente en R4Database")

btn_start = tk.Button(root, text="Iniciar Carga y Validación",
                    command=lambda: threading.Thread(target=ejecutar_proceso).start(),
                    font=("Segoe UI", 11, "bold"), bg="#4CAF50", fg="white")
btn_start.pack(pady=10)

root.mainloop()