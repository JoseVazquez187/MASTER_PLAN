import sqlite3
import pandas as pd
from tqdm import tqdm
import os
from tkinter import Tk, filedialog
# === CONFIGURACIÓN DE RUTAS ===

base_username = os.getlogin()
desktop_path = os.path.join("C:\\Users", base_username, "Desktop")

# === Selección del archivo de Expedite ===
Tk().withdraw()
path_exp_db = filedialog.askopenfilename(
    title="Selecciona el Expedite_final.db",
    filetypes=[("SQLite DB", "*.db")],
    initialdir=os.path.join(desktop_path, "PARCHE EXPEDITE")
)
if not path_exp_db:
    raise SystemExit("❌ No se seleccionó ningún archivo de Expedite.")

# === Validar y crear ruta ComprasDB\DataBase ===
compras_db_folder = os.path.join(desktop_path, "ComprasDB", "DataBase")
if not os.path.exists(compras_db_folder):
    os.makedirs(compras_db_folder)
path_oh_db = os.path.join(compras_db_folder, "compras_DB.db")

# === Validar y crear ruta Coverage ===
coverage_folder = os.path.join(desktop_path, "Coverage")
if not os.path.exists(coverage_folder):
    os.makedirs(coverage_folder)
output_db = os.path.join(coverage_folder, "CoberturaMateriales.db")

# === Obtener OH ===
def get_oh_information():
    conn = sqlite3.connect(path_oh_db)
    oh = pd.read_sql_query("SELECT * FROM oh", conn)
    exception_items = pd.read_sql_query("SELECT * FROM ExceptionItems", conn)
    conn.close()
    oh['OH'] = oh['OH'].astype(float)
    oh['ItemNo'] = oh['ItemNo'].str.upper()
    exception_items['ItemNo'] = exception_items['ItemNo'].str.upper()
    oh = pd.merge(oh, exception_items, how='left', on='ItemNo').fillna('')
    for index, row in oh.iterrows():
        if row['Replace'] != '':
            oh.at[index, 'ItemNo'] = row['Replace']
    return oh.groupby('ItemNo', as_index=False)['OH'].sum()

# === Obtener WO ===
def get_wo_information():
    conn = sqlite3.connect(path_oh_db)
    wo = pd.read_sql_query("SELECT * FROM WOInquiry", conn)
    conn.close()
    wo['ItemNumber'] = wo['ItemNumber'].str.upper().str.strip()
    wo = wo.rename(columns={'ItemNumber': 'ItemNo'})
    wo['OpnQ'] = wo['OpnQ'].astype(float)
    return wo[['ItemNo', 'WONo', 'OpnQ']]

# === Obtener Rework por lote ===
def get_rework_information_lote():
    conn = sqlite3.connect(path_oh_db)
    rwk = pd.read_sql_query("SELECT * FROM ReworkLoc_all", conn)
    conn.close()
    rwk = rwk.fillna('')
    rwk['ItemNo'] = rwk['ItemNo'].str.upper().str.strip()
    rwk['OH'] = rwk['OH'].astype(float)
    rwk['Lot'] = rwk['Lot'].astype(str).str.strip()
    return rwk.sort_values(by=['ItemNo', 'Expire_Date'])

# === Obtener Open Orders ===
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

# === OBTENER DINÁMICAMENTE TODAS LAS COLUMNAS DEL EXPEDITE ===
print("🔍 Detectando columnas del expedite...")
conn_temp = sqlite3.connect(path_exp_db)
expedite_columns_info = pd.read_sql("PRAGMA table_info(expedite)", conn_temp)
expedite_columns = expedite_columns_info['name'].tolist()
conn_temp.close()

print(f"✅ Detectadas {len(expedite_columns)} columnas en expedite:")
for i, col in enumerate(expedite_columns, 1):
    print(f"  {i:2d}. {col}")

# === Crear DB vacía antes de insertar progresivamente ===
if os.path.exists(output_db):
    os.remove(output_db)
conn_out = sqlite3.connect(output_db)
cur_out = conn_out.cursor()

# === CREAR TABLA DINÁMICAMENTE ===
# Columnas de cobertura (fijas)
coverage_columns = [
    "ItemNo TEXT", "ReqDate TEXT", "ReqQty REAL", "Index_Item INTEGER",
    "OH_Coverage REAL", "OpnQ REAL", "OpnQ_PO REAL", "Fill TEXT",
    "WO_used TEXT", "FillLot TEXT", "PO_used TEXT", "Balance REAL"
]

# Columnas que ya están en cobertura (para evitar duplicados)
coverage_column_names = ["ItemNo", "ReqDate", "ReqQty"]

# Columnas del expedite (dinámicas) - excluir las que ya están en cobertura
expedite_table_columns = []
for col in expedite_columns:
    if col in coverage_column_names:
        continue  # Saltar columnas que ya están en cobertura
    elif col == 'OH':
        expedite_table_columns.append("OH_Original REAL")
    else:
        # Escapar nombres de columnas con caracteres especiales
        col_escaped = f'"{col}"' if '-' in col or ' ' in col else col
        expedite_table_columns.append(f"{col_escaped} TEXT")

all_columns = coverage_columns + expedite_table_columns
create_table_sql = f"""
CREATE TABLE IF NOT EXISTS Cobertura (
    {', '.join(all_columns)}
)
"""

cur_out.execute(create_table_sql)
cur_out.execute("""
CREATE TABLE IF NOT EXISTS CoberturaPOs (
    ItemNo TEXT, PO_Line TEXT, QtyUsed REAL,
    RevPrDt TEXT, ReqDate TEXT
)
""")
conn_out.commit()

# === CARGAR TODAS LAS COLUMNAS DINÁMICAMENTE ===
print("📊 Cargando datos del expedite...")
conn = sqlite3.connect(path_exp_db)
# Construir SELECT dinámicamente
columns_for_select = []
for col in expedite_columns:
    if '-' in col or ' ' in col:
        columns_for_select.append(f'"{col}"')
    else:
        columns_for_select.append(col)

select_sql = f"SELECT {', '.join(columns_for_select)} FROM expedite"
df = pd.read_sql(select_sql, conn)
conn.close()

# === PREPARAR AGRUPAMIENTO DINÁMICO ===
df['ReqDate'] = pd.to_datetime(df['ReqDate'])
df['ReqQty'] = pd.to_numeric(df['ReqQty'], errors='coerce')

# Crear diccionario de agregación dinámicamente
agg_dict = {'ReqQty': 'sum'}
for col in expedite_columns:
    if col not in ['ItemNo', 'ReqDate', 'ReqQty']:
        agg_dict[col] = 'first'

df_grouped = df.groupby(['ItemNo', 'ReqDate'], as_index=False).agg(agg_dict)
df_grouped = df_grouped.sort_values(by=['ItemNo', 'ReqDate'])
df_grouped['Index_Item'] = df_grouped.groupby('ItemNo').cumcount() + 1
df_grouped['ItemNo'] = df_grouped['ItemNo'].astype(str).str.upper()

# === Obtener fuentes de cobertura ===
print("🔍 Obteniendo fuentes de cobertura...")
piv_oh = get_oh_information()
df_wo_raw = get_wo_information()
piv_rwk = get_rework_information_lote()
piv_po_raw = get_open_order_information()

# === Preparar estructuras ===
piv_wo_sum = df_wo_raw.groupby('ItemNo', as_index=False)['OpnQ'].sum()
piv_po_sum = piv_po_raw.groupby('ItemNo', as_index=False)['OpnQ'].sum()

# Renombrar columna OH del piv_oh para evitar conflicto con OH del expedite
piv_oh = piv_oh.rename(columns={'OH': 'OH_Coverage'})

df_cov = df_grouped.merge(piv_oh, on='ItemNo', how='left') \
                .merge(piv_wo_sum, on='ItemNo', how='left') \
                .merge(piv_po_sum, on='ItemNo', how='left', suffixes=('', '_PO'))

cols = ['OH_Coverage', 'OpnQ', 'OpnQ_PO']
df_cov[cols] = df_cov[cols].apply(pd.to_numeric, errors='coerce').fillna(0).astype(float)

# === Agrupar por ItemNo para insertar por ítem ===
rwk_dict = {k: g.copy() for k, g in piv_rwk.groupby('ItemNo')}
wo_dict = {k: g.copy() for k, g in df_wo_raw.groupby('ItemNo')}
po_dict = {k: g.copy() for k, g in piv_po_raw.groupby('ItemNo')}

print("🚀 Procesando cobertura por ítem...")
for item, group in tqdm(df_cov.groupby('ItemNo'), desc="Insertando por ítem"):
    group = group.sort_values(by='ReqDate')
    
    estado = {
        'oh': group['OH_Coverage'].iloc[0],
        'wo': wo_dict.get(item, pd.DataFrame()),
        'rwk': rwk_dict.get(item, pd.DataFrame()),
        'po': po_dict.get(item, pd.DataFrame())
    }

    for _, row in group.iterrows():
        restante = float(row['ReqQty'])
        req_date = row['ReqDate']
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
                                (item, po_line, usado, str(fecha_llegada), str(req_date))
                            )

                    if restante == 0:
                        fill = 'PO'
                        balance = estado['po']['OpnQ'].sum()
                    else:
                        fill = '❌ Faltante'
                        balance = -restante

        # === INSERTAR DINÁMICAMENTE ===
        # Valores de cobertura (fijos)
        coverage_values = [
            item, str(req_date), row['ReqQty'], index_item,
            row['OH_Coverage'], row['OpnQ'], row['OpnQ_PO'],
            fill, ', '.join(wo_used), ', '.join(lotes_usados), 
            ', '.join(po_usados), balance
        ]
        
        # Valores del expedite (dinámicos) - excluir columnas que ya están en cobertura
        expedite_values = []
        for col in expedite_columns:
            if col in coverage_column_names:
                continue  # Saltar columnas que ya están en cobertura
            elif col == 'OH':
                # OH del expedite se convierte en OH_Original
                value = row.get('OH', 0)
                if pd.notna(value) and col in ['OH', 'ReqQty'] or str(value).replace('.','').replace('-','').isdigit():
                    expedite_values.append(float(value))
                else:
                    expedite_values.append(0.0)
            else:
                expedite_values.append(str(row.get(col, '')))
        
        all_values = coverage_values + expedite_values
        
        # Crear placeholders dinámicamente
        placeholders = ', '.join(['?'] * len(all_values))
        
        # Crear lista de columnas para INSERT
        coverage_col_names = [
            "ItemNo", "ReqDate", "ReqQty", "Index_Item", "OH_Coverage", "OpnQ", "OpnQ_PO",
            "Fill", "WO_used", "FillLot", "PO_used", "Balance"
        ]
        
        expedite_col_names = []
        for col in expedite_columns:
            if col in coverage_column_names:
                continue  # Saltar columnas que ya están en cobertura
            elif col == 'OH':
                expedite_col_names.append('OH_Original')
            elif '-' in col or ' ' in col:
                expedite_col_names.append(f'"{col}"')
            else:
                expedite_col_names.append(col)
        
        all_column_names = coverage_col_names + expedite_col_names
        columns_str = ', '.join(all_column_names)
        
        insert_sql = f"INSERT INTO Cobertura ({columns_str}) VALUES ({placeholders})"
        cur_out.execute(insert_sql, all_values)

    conn_out.commit()

conn_out.close()
print(f"✅ Proceso completado. Base de datos creada con {len(expedite_columns)} columnas del expedite.")