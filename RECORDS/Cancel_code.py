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

        df['REF'] = df['REF'].astype(str).str.strip().str.replace("'", "")
        df['SUB'] = df['SUB'].astype(str).str.strip().str.replace("'", "")
        df['PO_Linea'] = df['REF'] + '-' + df['SUB']
        df['A-CD'] = df[a_cd_column].astype(str).str.upper().str.strip()
        df['Fecha_Archivo'] = fecha_archivo.strftime("%Y-%m-%d")

        df_filtrado = df[df['A-CD'].isin(codigos_seleccionados)][['PO_Linea', 'A-CD', 'Fecha_Archivo']]
        if not df_filtrado.empty:
            df_filtrado['PO_Linea'] = df_filtrado['PO_Linea'].str.replace("'", "")
            df_filtrado.to_sql("resultado_actions", conn, if_exists="append", index=False)

        cursor = conn.cursor()
        cursor.execute("INSERT INTO registro_archivos (archivo, fecha) VALUES (?, ?)", (archivo_normalizado, timestamp))
        conn.commit()

        return f"✅ {archivo} insertado con {len(df_filtrado)} registros."

    except Exception as e:
        return f"❌ Error procesando {archivo}: {e}"

# === Crear tabla PR561 ===
def create_pr_5_61(db_path):
    conn = None
    try:
        name_file = 'pr 5 61.txt'
        carpeta_destino = os.path.dirname(db_path)
        input_file = os.path.join(carpeta_destino, name_file)

        if not os.path.exists(input_file):
            log_insert("⚠️ No se encontró el archivo 'pr 5 61.txt'. Se omite carga de PR561.")
            return

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute("DROP TABLE IF EXISTS pr561")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS pr561(
                id INTEGER PRIMARY KEY,
                Entity TEXT, Project TEXT, Component TEXT, FuseNo TEXT,
                Description TEXT, PlnType TEXT, Srt TEXT,
                QtyOh TEXT, QtyIssue TEXT, QtyPending TEXT,
                ReqQty TEXT, WoNo TEXT
            )
        """)
        conn.commit()

        df = pd.read_fwf(input_file, widths=[2,10,7,32,14,30,9,3,16,13,12,14,12,12,14,9], header=3)
        df.drop([0], axis=0, inplace=True)
        df = df.loc[~df['In'].isin(['**', '', 'TO', 'En'])]
        df = df.fillna('')

        df = df.rename(columns={
            'Fuse-No': 'FuseNo',
            'Component Description': 'Description',
            'Qty-Oh': 'QtyOh',
            'Qty-Issue': 'QtyIssue',
            'Qty-no-Iss': 'QtyPending',
            'Qty-Required': 'ReqQty',
            'Wo-No': 'WoNo'
        })

        df = df[['Entity', 'Project', 'Component', 'FuseNo', 'Description',
                'PlnType', 'Srt', 'QtyOh', 'QtyIssue', 'QtyPending', 'ReqQty', 'WoNo']]

        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].str.strip()

        df.to_sql('pr561', conn, if_exists='append', index=False)
        conn.commit()
        log_insert("✅ Se cargó el PR 5 61 correctamente en la tabla 'pr561'")
    except Exception as e:
        log_insert(f"❌ Error cargando PR 5 61: {e}")
    finally:
        if conn:
            conn.close()

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

    carpeta_destino = os.path.dirname(archivo_cancelaciones)
    db_path = os.path.join(carpeta_destino, "historico_actions.db")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    log_insert("🟡 Selecciona la carpeta del histórico de actions")
    path_historico = filedialog.askdirectory(title="Selecciona carpeta del histórico de actions")
    if not path_historico:
        log_insert("❌ No se seleccionó carpeta de histórico. Cancelando.")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS resultado_actions (PO_Linea TEXT, 'A-CD' TEXT, Fecha_Archivo TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS registro_archivos (archivo TEXT PRIMARY KEY, fecha TEXT)")
    conn.commit()

    archivos = sorted(os.listdir(path_historico))
    cursor.execute("SELECT archivo FROM registro_archivos")
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

        df_cancelaciones['REF'] = df_cancelaciones['REF'].astype(str).str.strip().str.replace("'", "")
        df_cancelaciones['SUB'] = df_cancelaciones['SUB'].astype(str).str.strip().str.replace("'", "")
        df_cancelaciones['PO_Linea'] = df_cancelaciones['REF'] + '-' + df_cancelaciones['SUB']
        df_cancelaciones['A-CD'] = df_cancelaciones['A-CD'].astype(str).str.upper().str.strip()
        df_cancelaciones = df_cancelaciones[df_cancelaciones['A-CD'].isin(codigos_seleccionados)]

        query = """
            SELECT PO_Linea, "A-CD", MIN(Fecha_Archivo) AS Fecha_Primera_Aparicion
            FROM resultado_actions
            GROUP BY PO_Linea, "A-CD"
        """
        df_sqlite = pd.read_sql_query(query, conn)

        df_resultado = df_cancelaciones.merge(df_sqlite, on=['PO_Linea', 'A-CD'], how='left')
        df_resultado['Fecha_Primera_Aparicion'] = pd.to_datetime(df_resultado['Fecha_Primera_Aparicion'], errors='coerce')
        df_resultado['Días_Sin_Ejecutar'] = (datetime.today() - df_resultado['Fecha_Primera_Aparicion']).dt.days
        
        
        
        
        df_resultado.to_sql("resultado_final", conn, if_exists="replace", index=False)
        log_insert("📁 Resultado final insertado en la tabla 'resultado_final' de la base de datos.")
    except Exception as e:
        log_insert(f"❌ Error generando resultado final: {e}")

    conn.commit()
    conn.close()
    create_pr_5_61(db_path)

btn_start = tk.Button(root, text="Iniciar Carga y Validación",
                    command=lambda: threading.Thread(target=ejecutar_proceso).start(),
                    font=("Segoe UI", 11, "bold"), bg="#4CAF50", fg="white")
btn_start.pack(pady=10)

root.mainloop()
