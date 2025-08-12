import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import os
from tkinter import filedialog, Tk

# === Selección automática del usuario ===
base_username = os.getlogin()
# username_with_one = base_username + ".ONE"

# Verificar si existe la cuenta .ONE
# user_path = os.path.join("C:\\Users", username_with_one) if os.path.exists(os.path.join("C:\\Users", username_with_one)) else os.path.join("C:\\Users", base_username)
user_path = "C:\\Users\\j.vazquez"
username = os.path.basename(user_path)

# === Ruta del Expedite.csv ===
expedite_path = os.path.join(user_path, "Desktop", "PARCHE EXPEDITE", "Expedite.csv")
if not os.path.exists(expedite_path):
    print(f"❌ No se encontró el archivo Expedite en: {expedite_path}")
    exit()

# === Seleccionar base SQLite ===
Tk().withdraw()
db_path = filedialog.askopenfilename(title="Selecciona la Base_3 y la fecha de hoy", filetypes=[("SQLite DB", "*.db")])
if not db_path:
    print("❌ No se seleccionó base.")
    exit()

# === Leer lt_total ===
conn_lt = sqlite3.connect(db_path)
df_lt = pd.read_sql("SELECT * FROM lt_total", conn_lt)
conn_lt.close()

df_lt[['Component', 'Sort', 'ItemNo']] = df_lt[['Component', 'Sort', 'ItemNo']].fillna('')

df_lt["KEY_UNICA"] = (
    df_lt["Component"].astype(str).str.strip().str.upper() + "_" +
    df_lt["Sort"].astype(str).str.strip().str.upper() + "_" +
    df_lt["ItemNo"].astype(str).str.strip().str.upper()
)


df_lt = df_lt[["KEY_UNICA", "LT_TOTAL"]]
df_lt["LT_TOTAL"] = pd.to_numeric(df_lt["LT_TOTAL"], errors="coerce")
df_lt = df_lt.groupby("KEY_UNICA", as_index=False).agg({"LT_TOTAL": "max"})

# === Leer Expedite ===
df_exp = pd.read_csv(expedite_path,skiprows = 1)
df_exp.columns = df_exp.columns.str.strip()

column_map = {
    'Entity Group': 'EntityGroup',
    'A/C': 'AC',
    'Item-No': 'ItemNo',
    'Req-Qty': 'ReqQty',
    'Demand-Source': 'DemandSource',
    'Req-Date': 'ReqDate',
    'Ship-Date': 'ShipDate',
    'MLIK Code': 'MLIKCode',
    'STD-Cost': 'STDCost',
    'Lot-Size': 'LotSize',
    'Fill-Doc': 'FillDoc',
    'Demand - Type': 'DemandType'
}
df_exp = df_exp.rename(columns={k: v for k, v in column_map.items() if k in df_exp.columns})

# === Crear KEY_UNICA y hacer merge ===

df_exp[['ItemNo', 'Sort', 'DemandSource']] = df_exp[['ItemNo', 'Sort', 'DemandSource']].fillna('')
df_exp["KEY_UNICA"] = (
    df_exp["ItemNo"].astype(str).str.strip() + "_" +
    df_exp["Sort"].astype(str).str.strip() + "_" +
    df_exp["DemandSource"].astype(str).str.strip()
).str.upper()


df_merged = pd.merge(df_exp, df_lt, on="KEY_UNICA", how="left")

# === Renombrar ReqDate y calcular nueva ===
df_merged["ReqDate_Original"] = df_merged["ReqDate"]
df_merged["ReqDate"] = pd.to_datetime(df_merged["ShipDate"], errors="coerce")

df_merged["ReqDate"] = df_merged.apply(
    lambda row: row["ReqDate"] - timedelta(days=int(row["LT_TOTAL"])) if pd.notna(row["ReqDate"]) and pd.notna(row["LT_TOTAL"]) else row["ReqDate"],
    axis=1
)

# Si DemandType es SAFE y LT_TOTAL está vacío, asignar 2
df_merged.loc[
    (df_merged['DemandType'].astype(str).str.upper() == 'SAFE') & (df_merged['LT_TOTAL'].isna()),
    'LT_TOTAL'
] = 2

# === Guardar Expedite_final.db ===
# === Guardar Expedite_final.db ===
output_path = os.path.join(os.path.dirname(db_path), "Expedite_final.db")
conn_out = sqlite3.connect(output_path)
df_merged.to_sql("expedite", conn_out, if_exists="replace", index=False)
conn_out.close()

# === Crear base adicional BS_Analisis_fecha.db ===
today_str = datetime.today().strftime("%Y_%m_%d")
bs_analisis_path = os.path.join(os.path.dirname(db_path), f"BS_Analisis_{today_str}.db")

conn_bs = sqlite3.connect(bs_analisis_path)
df_merged.to_sql("expedite", conn_bs, if_exists="replace", index=False)
conn_bs.close()

# === Validación de integridad: Expedite_final.db ===
conn_check = sqlite3.connect(output_path)
total_en_csv = len(df_merged)
total_en_sql = pd.read_sql_query("SELECT COUNT(*) as total FROM expedite", conn_check).iloc[0]['total']
conn_check.close()

if total_en_csv == total_en_sql:
    print(f"✅ Validación completada: Expedite_final.db contiene {total_en_csv} registros.")
else:
    print(f"❌ Advertencia: Expedite_final.db no coincide.")
    print(f"   - Registros esperados: {total_en_csv}")
    print(f"   - Registros reales: {total_en_sql}")

# === Validación de integridad: BS_Analisis_YYYY_MM_DD.db ===
conn_check_bs = sqlite3.connect(bs_analisis_path)
total_en_sql_bs = pd.read_sql_query("SELECT COUNT(*) as total FROM expedite", conn_check_bs).iloc[0]['total']
conn_check_bs.close()

if total_en_csv == total_en_sql_bs:
    print(f"✅ Validación completada: BS_Analisis_{today_str}.db contiene {total_en_sql_bs} registros.")
else:
    print(f"❌ Advertencia: BS_Analisis_{today_str}.db no coincide.")
    print(f"   - Registros esperados: {total_en_csv}")
    print(f"   - Registros reales: {total_en_sql_bs}")

# === Confirmaciones finales ===
print(f"✅ Expedite_final.db creada correctamente en:\n{output_path}")
print(f"✅ BS_Analisis_{today_str}.db creada correctamente en:\n{bs_analisis_path}")
