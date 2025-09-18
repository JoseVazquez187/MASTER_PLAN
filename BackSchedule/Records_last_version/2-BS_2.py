import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import os
import tkinter as tk
from tkinter import filedialog

# === Feriados definidos ===
holiday_dates = pd.to_datetime([
    "2023-01-01", "2023-02-06", "2023-03-20", "2023-04-07", "2023-05-01", "2023-09-16", "2023-11-20",
    "2023-12-24", "2023-12-25", "2023-12-26", "2023-12-27", "2023-12-28", "2023-12-29", "2023-12-30", "2023-12-31",
    "2024-01-01", "2024-02-05", "2024-03-18", "2024-03-29", "2024-05-01", "2024-09-16", "2024-10-01", "2024-11-18",
    "2024-12-23", "2024-12-24", "2024-12-25", "2024-12-26", "2024-12-27", "2024-12-30",
    "2025-01-01", "2025-02-03", "2025-03-17", "2025-04-17", "2025-05-01", "2025-09-15", "2025-11-17",
    "2025-12-25", "2025-12-26", "2025-12-29", "2025-12-30", "2025-12-31",
    "2026-01-01"
]).date

# === Función para contar días hábiles ===
def contar_dias_habiles(fecha_inicio, fecha_fin, feriados):
    dias = 0
    fecha_actual = fecha_inicio
    while fecha_actual < fecha_fin:
        if fecha_actual.weekday() < 5 and fecha_actual.date() not in feriados:
            dias += 1
        fecha_actual += timedelta(days=1)
    return dias

# === Selector de archivo SQLite ===
root = tk.Tk()
root.withdraw()
path_db = filedialog.askopenfilename(
    title="Selecciona la base de datos con Base1 y la fecha de hoy",
    filetypes=[("SQLite DB", "*.db")]
)
if not path_db:
    print("❌ No se seleccionó base de datos.")
    exit()

# === Cargar base de datos seleccionada ===
conn = sqlite3.connect(path_db)
df = pd.read_sql("SELECT * FROM programacion_materiales", conn)
df['ShipDate'] = pd.to_datetime(df['ShipDate'], errors='coerce')
df['BackScheduleDate'] = pd.to_datetime(df['BackScheduleDate'], errors='coerce')

# === Crear diccionario de plantypes desde archivo externo si existe, si no desde la misma base ===
path_plantype_excel = os.path.join(os.path.dirname(path_db), "PlanTypes_areas_primarias.xlsx")
if os.path.exists(path_plantype_excel):
    df_pt = pd.read_excel(path_plantype_excel)
    plantype_dict = dict(zip(df_pt['Plan-Type'].astype(str).str.strip().str.upper(), df_pt['LT_PlanType']))
else:
    df_pt = df[['Plan-Type', 'LT_PlanType']].dropna().drop_duplicates()
    plantype_dict = dict(zip(df_pt['Plan-Type'].astype(str).str.strip().str.upper(), df_pt['LT_PlanType']))

# === Normalización y preparación de columnas ===
df['Plan-Type'] = df['Plan-Type'].astype(str).str.strip().str.upper()
df['LT_PlanType'] = None
df['LT_Real'] = None
df['LT_Remanente_Usable'] = None
df['LT_TOTAL'] = None

# === Aplicar lógica de remanente por key solo si Plan-Type es válido ===
for key, grupo in df.groupby("key"):
    for idx, row in grupo.iterrows():
        pt = row['Plan-Type']
        if pt in plantype_dict and pd.notna(row['BackScheduleDate']) and pd.notna(row['ShipDate']):
            lt_plan = plantype_dict[pt]
            lt_real = contar_dias_habiles(row['BackScheduleDate'], row['ShipDate'], holiday_dates)
            lt_remanente = lt_plan - lt_real
            indices = df[df['key'] == key].index
            df.loc[indices, 'LT_PlanType'] = lt_plan
            df.loc[indices, 'LT_Real'] = lt_real
            df.loc[indices, 'LT_Remanente_Usable'] = lt_remanente
            break  # Solo se usa el primer match del grupo

# === Calcular LT_TOTAL ===
df['LT_TOTAL'] = df.apply(lambda r: r['OFFSET_DIAS'] + r['LT_Remanente_Usable']
                        if pd.notna(r['OFFSET_DIAS']) and pd.notna(r['LT_Remanente_Usable']) else None, axis=1)

# === Guardar nuevo resultado en una base nueva ===
nombre_resultado = "Base2_" + datetime.now().strftime("%Y_%m_%d_%H_%M") + ".db"
path_resultado = os.path.join(os.path.dirname(path_db), nombre_resultado)
conn_out = sqlite3.connect(path_resultado)
df.to_sql("lt_total", conn_out, if_exists='replace', index=False)
conn_out.close()

print(f"✅ Archivo generado: {path_resultado}")
