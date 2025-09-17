import sqlite3
import pandas as pd
from tqdm import tqdm
import os
from tkinter import Tk, filedialog
from datetime import datetime, timedelta
import getpass
import numpy as np 
# === CONFIGURACIÓN DE RUTAS ===
base_username = getpass.getuser()
desktop_path = os.path.join("C:\\Users", base_username, "Desktop")

Tk().withdraw()
path_exp_db = filedialog.askopenfilename(
    title="Selecciona el Expedite_final.db",
    filetypes=[("SQLite DB", "*.db")],
    initialdir=os.path.join(desktop_path, "PARCHE EXPEDITE")
)
if not path_exp_db:
    raise SystemExit("❌ No se seleccionó ningún archivo de Expedite.")

compras_db_folder = os.path.join(desktop_path, "ComprasDB", "DataBase")
os.makedirs(compras_db_folder, exist_ok=True)
path_oh_db = os.path.join(compras_db_folder, "compras_DB.db")

coverage_folder = os.path.join(desktop_path, "Coverage")
os.makedirs(coverage_folder, exist_ok=True)
output_db = os.path.join(coverage_folder, "CoberturaMaterialesXweek.db")

# === FUNCIONES PARA OBTENER DATOS ===
def get_oh_information():
    conn = sqlite3.connect(path_oh_db)
    oh = pd.read_sql_query("SELECT * FROM oh", conn)
    exception_items = pd.read_sql_query("SELECT * FROM ExceptionItems", conn)
    conn.close()
    oh['OH'] = oh['OH'].astype(float)
    oh['ItemNo'] = oh['ItemNo'].str.upper()
    exception_items['ItemNo'] = exception_items['ItemNo'].str.upper()
    oh = pd.merge(oh, exception_items, how='left', on='ItemNo').fillna('')
    oh.loc[oh['Replace'] != '', 'ItemNo'] = oh['Replace']
    return oh.groupby('ItemNo', as_index=False)['OH'].sum()

def get_wo_information():
    conn = sqlite3.connect(path_oh_db)
    wo = pd.read_sql_query("SELECT * FROM WOInquiry", conn)
    conn.close()
    wo['ItemNumber'] = wo['ItemNumber'].str.upper().str.strip()
    wo = wo.rename(columns={'ItemNumber': 'ItemNo'})
    wo['OpnQ'] = wo['OpnQ'].astype(float)
    return wo[['ItemNo', 'WONo', 'OpnQ']]

def get_rework_information_lote():
    conn = sqlite3.connect(path_oh_db)
    rwk = pd.read_sql_query("SELECT * FROM ReworkLoc_all", conn)
    conn.close()
    rwk = rwk.fillna('')
    rwk['ItemNo'] = rwk['ItemNo'].str.upper().str.strip()
    rwk['OH'] = rwk['OH'].astype(float)
    rwk['Lot'] = rwk['Lot'].astype(str).str.strip()
    return rwk.sort_values(by=['ItemNo', 'Expire_Date'])

def get_open_order_information():
    conn = sqlite3.connect(path_oh_db)
    oo = pd.read_sql_query("SELECT * FROM openOrder", conn)
    conn.close()
    oo = oo.fillna('')
    oo['ItemNo'] = oo['ItemNo'].str.upper().str.strip()
    oo['PO_Line'] = oo['PONo'].astype(str).str.strip() + '-' + oo['Ln'].astype(str).str.strip()
    oo['OpnQ'] = oo['OpnQ'].astype(float)
    oo['RevPrDt'] = pd.to_datetime(oo['RevPrDt'])
    return oo.sort_values(by=['ItemNo', 'RevPrDt'])

# === INICIALIZA Y CREA BASE DE DATOS VACÍA ===
if os.path.exists(output_db):
    os.remove(output_db)
conn_out = sqlite3.connect(output_db)
cur_out = conn_out.cursor()

cur_out.execute("""
CREATE TABLE IF NOT EXISTS Cobertura (
    ItemNo TEXT, Semana TEXT, ReqQty REAL, Index_Item INTEGER,
    OH REAL, OpnQ REAL, OpnQ_PO REAL, Fill TEXT,
    WO_used TEXT, FillLot TEXT, PO_used TEXT, Balance REAL,
    Cobertura_Acumulada_Porc REAL
)
""")
cur_out.execute("""
CREATE TABLE IF NOT EXISTS CoberturaPOs (
    ItemNo TEXT, PO_Line TEXT, QtyUsed REAL,
    RevPrDt TEXT, Semana TEXT
)
""")
conn_out.commit()

# === CARGA EXPEDITE Y LIMPIEZA ===
conn = sqlite3.connect(path_exp_db)
df = pd.read_sql("SELECT ItemNo, ReqDate, ReqQty FROM expedite", conn)
conn.close()

df['ReqDate'] = pd.to_datetime(df['ReqDate'], errors='coerce')
df = df.dropna(subset=['ReqDate'])
df['ReqQty'] = pd.to_numeric(df['ReqQty'], errors='coerce')
df['ItemNo'] = df['ItemNo'].str.upper()

# === GENERAR SEMANA Y SEMANA_ORDEN ===
today = pd.Timestamp.today().normalize()

def get_semana(row):
    if row['ReqDate'] < today:
        return 'PAST DUE'
    else:
        sem = row['ReqDate'].isocalendar()
        return f"W{sem.week:02d}-{sem.year}"

def get_semana_orden(row):
    if row['ReqDate'] < today:
        return -1
    else:
        sem = row['ReqDate'].isocalendar()
        return sem.year * 100 + sem.week

df['Semana'] = df.apply(get_semana, axis=1)
df['Semana_Orden'] = df.apply(get_semana_orden, axis=1)

# === AGRUPA POR SEMANA ===
df_grouped = df.groupby(['ItemNo', 'Semana', 'Semana_Orden'], as_index=False)['ReqQty'].sum()
df_grouped = df_grouped.sort_values(by=['ItemNo', 'Semana_Orden'])
df_grouped['Index_Item'] = df_grouped.groupby('ItemNo').cumcount() + 1

# === CARGA INVENTARIO Y PEDIDOS ===
piv_oh = get_oh_information()
df_wo_raw = get_wo_information()
piv_rwk = get_rework_information_lote()
df_po_raw = get_open_order_information()

piv_wo_sum = df_wo_raw.groupby('ItemNo', as_index=False)['OpnQ'].sum()
piv_po_sum = df_po_raw.groupby('ItemNo', as_index=False)['OpnQ'].sum()
df_cov = df_grouped.merge(piv_oh, on='ItemNo', how='left') \
                .merge(piv_wo_sum, on='ItemNo', how='left') \
                .merge(piv_po_sum, on='ItemNo', how='left', suffixes=('', '_PO'))

for col in ['OH', 'OpnQ', 'OpnQ_PO']:
    df_cov[col] = pd.to_numeric(df_cov[col], errors='coerce').fillna(0).astype(float)

# === ASIGNA INVENTARIOS, WO, RWK Y PO ===
rwk_dict = {k: g.copy() for k, g in piv_rwk.groupby('ItemNo')}
wo_dict = {k: g.copy() for k, g in df_wo_raw.groupby('ItemNo')}
po_dict = {k: g.copy() for k, g in df_po_raw.groupby('ItemNo')}

for item, group in tqdm(df_cov.groupby('ItemNo')):
    group = group.sort_values(by='Semana_Orden')
    estado = {
        'oh': group['OH'].iloc[0],
        'wo': wo_dict.get(item, pd.DataFrame()),
        'rwk': rwk_dict.get(item, pd.DataFrame()),
        'po': po_dict.get(item, pd.DataFrame())
    }
    for _, row in group.iterrows():
        semana = row['Semana']
        restante = float(row['ReqQty'])
        index_item = int(row['Index_Item'])
        wo_used, lotes_usados, po_usados = [], [], []
        fill = ''
        balance = 0

        if estado['oh'] >= restante:
            estado['oh'] -= restante
            fill = 'OH'
            balance = estado['oh']
        else:
            restante -= estado['oh']
            estado['oh'] = 0

            for j in estado['wo'].index:
                if restante <= 0:
                    break
                q = estado['wo'].at[j, 'OpnQ']
                if q > 0:
                    usado = min(restante, q)
                    restante -= usado
                    estado['wo'].at[j, 'OpnQ'] -= usado
                    wo_used.append(str(estado['wo'].at[j, 'WONo']))

            if restante == 0:
                fill = 'WO'
                balance = estado['wo']['OpnQ'].sum()
            else:
                for j in estado['rwk'].index:
                    if restante <= 0:
                        break
                    q = estado['rwk'].at[j, 'OH']
                    if q > 0:
                        usado = min(restante, q)
                        restante -= usado
                        estado['rwk'].at[j, 'OH'] -= usado
                        lotes_usados.append(estado['rwk'].at[j, 'Lot'])

                if restante == 0:
                    fill = 'Rework'
                    balance = estado['rwk']['OH'].sum()
                else:
                    for j in estado['po'].index:
                        if restante <= 0:
                            break
                        q = estado['po'].at[j, 'OpnQ']
                        if q > 0:
                            usado = min(restante, q)
                            restante -= usado
                            estado['po'].at[j, 'OpnQ'] -= usado
                            po_line = estado['po'].at[j, 'PO_Line']
                            fecha_llegada = estado['po'].at[j, 'RevPrDt']
                            po_usados.append(po_line)
                            cur_out.execute(
                                "INSERT INTO CoberturaPOs VALUES (?, ?, ?, ?, ?)",
                                (item, po_line, usado, str(fecha_llegada), semana)
                            )

                    if restante == 0:
                        fill = 'PO'
                        balance = estado['po']['OpnQ'].sum()
                    else:
                        fill = '❌ Faltante'
                        balance = -restante

        cur_out.execute("""
            INSERT INTO Cobertura (
                ItemNo, Semana, ReqQty, Index_Item, OH, OpnQ, OpnQ_PO,
                Fill, WO_used, FillLot, PO_used, Balance, Cobertura_Acumulada_Porc
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            item, semana, row['ReqQty'], index_item,
            row['OH'], row['OpnQ'], row['OpnQ_PO'],
            fill,
            ', '.join(wo_used),
            ', '.join(lotes_usados),
            ', '.join(po_usados),
            balance, 0.0
        ))

    # Cobertura acumulada
    cur_out.execute("SELECT rowid, Semana, ReqQty, Balance FROM Cobertura WHERE ItemNo = ? ORDER BY CASE Semana WHEN 'PAST DUE' THEN -1 ELSE CAST(substr(Semana, 2, 2) AS INTEGER) + CAST(substr(Semana, 6, 4)*100 AS INTEGER) END", (item,))
    rows = cur_out.fetchall()
    acumulado = 0
    total_requerido = sum(r[2] for r in rows)
    for r in rows:
        rowid, semana, reqqty, balance = r
        cantidad_cubierta = reqqty if balance >= 0 else (reqqty + balance)
        acumulado += max(0, cantidad_cubierta)
        porc = round((acumulado / total_requerido) * 100, 2) if total_requerido > 0 else 0
        cur_out.execute("UPDATE Cobertura SET Cobertura_Acumulada_Porc = ? WHERE rowid = ?", (porc, rowid))
    conn_out.commit()
conn_out.close()

# === AHORA PROCESA Y AGREGA SemanaLunes Y PO_Status EN PANDAS ===
with sqlite3.connect(output_db) as conn:
    df_cov = pd.read_sql("SELECT * FROM Cobertura", conn)
    df_pos = pd.read_sql("""
        SELECT ItemNo, Semana, MIN(PO_Line) AS PO_Line, MIN(RevPrDt) AS RevPrDt
        FROM CoberturaPOs
        GROUP BY ItemNo, Semana
    """, conn)

def primer_lunes_iso(semana):
    if pd.isna(semana):
        return None
    if semana == 'PAST DUE':
        hoy = datetime.today()
        lunes_pasado = hoy - timedelta(days=hoy.weekday() + 7)
        return lunes_pasado.date()
    try:
        week = int(semana[1:3])
        year = int(semana[4:8])
        return datetime.fromisocalendar(year, week, 1).date()
    except Exception:
        return None

df_cov['SemanaLunes'] = df_cov['Semana'].apply(primer_lunes_iso)

df_final = df_cov.merge(df_pos, how='left', on=['ItemNo', 'Semana'])
df_final['RevPrDt'] = pd.to_datetime(df_final['RevPrDt'], errors='coerce').dt.date

def po_status(row):
    if pd.isna(row['SemanaLunes']) or pd.isna(row['RevPrDt']):
        return 'N/A'
    if row['RevPrDt'] > row['SemanaLunes']:
        return 'PULL'
    elif row['RevPrDt'] == row['SemanaLunes']:
        return 'KEEP'
    elif row['RevPrDt'] < row['SemanaLunes']:
        return 'PUSH'
    else:
        return 'N/A'

df_final['PO_Status'] = df_final.apply(po_status, axis=1)

def semanas_action(row):
    if pd.isna(row['RevPrDt']) or pd.isna(row['SemanaLunes']):
        return None
    diff_days = (row['RevPrDt'] - row['SemanaLunes']).days
    semanas = int(np.ceil(diff_days / 7.0))
    if row['PO_Status'] == 'PULL':
        return semanas  # positivo
    elif row['PO_Status'] == 'PUSH':
        return semanas  # negativo si llega antes
    else:
        return 0

df_final['Semanas_Action'] = df_final.apply(semanas_action, axis=1)

with sqlite3.connect(output_db) as conn:
    df_final.to_sql('CoberturaAnalizada', conn, if_exists='replace', index=False)

print("✅ CoberturaAnalizada generada correctamente y lista para análisis.")
