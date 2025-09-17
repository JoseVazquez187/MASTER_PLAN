import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import os
from tkinter import filedialog, Tk
from tqdm import tqdm

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

def contar_dias_habiles(fecha_inicio, fecha_fin, feriados):
    dias = 0
    fecha_actual = fecha_inicio
    while fecha_actual < fecha_fin:
        if fecha_actual.weekday() < 5 and fecha_actual.date() not in feriados:
            dias += 1
        fecha_actual += timedelta(days=1)
    return dias

def contar_dias_totales(fecha_inicio, fecha_fin):
    return (fecha_fin - fecha_inicio).days

# === Seleccionar base ===
Tk().withdraw()
db_path = filedialog.askopenfilename(title="Selecciona la Base_2 con la fecha de hoy", filetypes=[("SQLite DB", "*.db")])
if not db_path:
    print("❌ No se seleccionó base.")
    exit()

# === Leer base ===
conn = sqlite3.connect(db_path)
df = pd.read_sql("SELECT * FROM lt_total", conn)
conn.close()

df['ShipDate'] = pd.to_datetime(df['ShipDate'], errors='coerce')
df['BackScheduleDate'] = pd.to_datetime(df['BackScheduleDate'], errors='coerce')
df['Plan-Type'] = df['Plan-Type'].astype(str).str.strip().str.upper()
df['Level Number'] = df['Level Number'].astype(str).str.strip()
df['Nivel'] = df['Level Number'].str.extract(r"(\d)").astype(float).fillna(0).astype(int)

# === Leer PlanTypes válidos ===
plan_path = os.path.join(os.path.dirname(db_path), "PlanTypes_areas_primarias.xlsx")
plan_df = pd.read_excel(plan_path)
valid_plantypes = set(plan_df['Plan-Type'].astype(str).str.strip().str.upper())

# === Salida ===
nombre_resultado = "Base3_" + datetime.now().strftime("%Y_%m_%d_%H_%M") + ".db"
path_resultado = os.path.join(os.path.dirname(db_path), nombre_resultado)
conn_out = sqlite3.connect(path_resultado)

# === Procesar por clave respetando orden original ===
df = df.sort_values(['key', 'Orden_BOM_Original']).reset_index(drop=True)
keys = df['key'].unique()

for key in tqdm(keys, desc="Procesando claves"):
    grupo = df[df['key'] == key].copy()
    grupo['LT_PlanType'] = None
    grupo['LT_Real'] = None
    grupo['LT_Remanente_Usable'] = None
    grupo['LT_TOTAL'] = None
    grupo['OFFSET_DIAS'] = grupo['OFFSET_DIAS']  # Mantener si ya existe

    calcular = False
    nivel_base = None
    lt_remanente = None
    lt_plan = None
    lt_real = None

    for idx in grupo.index:
        fila = grupo.loc[idx]
        pt = fila['Plan-Type']
        nivel = fila['Nivel']

        if nivel == 1 and pd.notna(fila['BackScheduleDate']) and pd.notna(fila['ShipDate']):
            grupo.at[idx, 'OFFSET_DIAS'] = contar_dias_totales(fila['BackScheduleDate'], fila['ShipDate'])
            continue

        if pt in valid_plantypes and pd.notna(fila['BackScheduleDate']) and pd.notna(fila['ShipDate']):
            try:
                lt_plan = float(plan_df.loc[plan_df['Plan-Type'].str.strip().str.upper() == pt, 'LT_PlanType'].values[0])
                lt_real = float(contar_dias_habiles(fila['BackScheduleDate'], fila['ShipDate'], holiday_dates))
                lt_remanente = lt_plan - lt_real
                calcular = True
                nivel_base = nivel

                grupo.at[idx, 'LT_PlanType'] = lt_plan
                grupo.at[idx, 'LT_Real'] = lt_real
                grupo.at[idx, 'LT_Remanente_Usable'] = lt_remanente

                # No colocar LT_TOTAL aquí (nivel con PlanType válido)
                continue
            except:
                calcular = False
                continue

        elif calcular and nivel > nivel_base:
            grupo.at[idx, 'LT_PlanType'] = lt_plan
            grupo.at[idx, 'LT_Real'] = lt_real
            grupo.at[idx, 'LT_Remanente_Usable'] = lt_remanente
            if pd.notna(fila['OFFSET_DIAS']):
                total = float(fila['OFFSET_DIAS']) + lt_remanente
                grupo.at[idx, 'LT_TOTAL'] = min(total, lt_plan)
        else:
            calcular = False
            nivel_base = None

    # ... for idx in grupo.index ... lógica previa ...

    # Asignar 2 si es nivel 1, tipo L, y sin LT_TOTAL
    grupo.loc[(grupo['Nivel'] == 1) & (grupo['MLI'].str.upper() == 'L') & (grupo['LT_TOTAL'].isna()), 'LT_TOTAL'] = 2
    grupo.loc[(grupo['Nivel'] == 1) & (grupo['MLI'].str.upper() == 'M') & (grupo['LT_TOTAL'].isna()), 'LT_TOTAL'] = 2
    

    # Fallback: usar LT_Real si LT_TOTAL está vacío
    # Primero: usar LT_Real si está disponible
    grupo.loc[grupo['LT_TOTAL'].isna() & grupo['LT_Real'].notna(), 'LT_TOTAL'] = grupo.loc[grupo['LT_TOTAL'].isna() & grupo['LT_Real'].notna(), 'LT_Real']
    # Luego: si sigue vacía, usar OFFSET_
    grupo.loc[grupo['LT_TOTAL'].isna() & grupo['OFFSET_DIAS'].notna(), 'LT_TOTAL'] = grupo.loc[grupo['LT_TOTAL'].isna() & grupo['OFFSET_DIAS'].notna(), 'OFFSET_DIAS']
    grupo.drop(columns=["Nivel"]).to_sql("lt_total", conn_out, if_exists='append', index=False)

conn_out.close()
print(f"✅ Base generada correctamente: {path_resultado}")
