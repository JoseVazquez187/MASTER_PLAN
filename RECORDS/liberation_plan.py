import pandas as pd
import sqlite3
from datetime import datetime, timedelta

# --- CONFIGURACIÓN ---
db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"

# --- 1. Cargar las tablas desde la base de datos ---
with sqlite3.connect(db_path) as conn:
    fcst = pd.read_sql("SELECT * FROM fcst", conn)
    wo = pd.read_sql("SELECT * FROM WOInquiry", conn)
    item = pd.read_sql("SELECT * FROM IN92", conn)

# --- 2. Normalizar columnas clave ---
fcst['ItemNo'] = fcst['Item-No'] if 'Item-No' in fcst.columns else fcst['ItemNo']
fcst['AC'] = fcst['AC-No'] if 'AC-No' in fcst.columns else fcst['AC']
fcst['ReqDate'] = pd.to_datetime(fcst['Reqt-Dt'] if 'Reqt-Dt' in fcst.columns else fcst['ReqDate'])

wo['ItemNo'] = wo['ItemNumber'] if 'ItemNumber' in wo.columns else wo['ItemNo']
wo['AC'] = wo['AC']
wo['DueDt'] = pd.to_datetime(wo['DueDt'], errors='coerce')
wo['Prtdate'] = pd.to_datetime(wo['Prtdate'], errors='coerce')
wo['key'] = wo['ItemNo'].astype(str) + '_' + wo['AC'].astype(str) + '_' + wo['DueDt'].dt.strftime('%Y-%m-%d')

item['ItemNo'] = item['ItemNo'] if 'ItemNo' in item.columns else item['Item-Number']
item['LT'] = pd.to_numeric(item['LT'], errors='coerce').fillna(0)

# --- 3. Merge para saber el LT de cada requerimiento ---
fcst = fcst.merge(item[['ItemNo', 'LT', 'Description']], on='ItemNo', how='left')

# --- 4. Calcular la fecha límite para liberar (ReqDate - LT días) ---
fcst['LT'] = fcst['LT'].fillna(0).astype(int)
fcst['ReleaseLimitDate'] = fcst['ReqDate'] - fcst['LT'].apply(lambda x: timedelta(days=x))

# --- 5. Llave para cruce: Item + AC + Fecha Req ---
fcst['key'] = fcst['ItemNo'].astype(str) + '_' + fcst['AC'].astype(str) + '_' + fcst['ReqDate'].dt.strftime('%Y-%m-%d')

# --- 6. Verificar si existe WO para esa necesidad ---
fcst['WO_Liberada'] = fcst['key'].isin(wo['key'])

# --- 7. Calcular estatus de liberación ---
today = pd.to_datetime(datetime.now().date())

def liberacion_status(row):
    if row['WO_Liberada']:
        return 'Liberado'
    elif row['ReleaseLimitDate'] < today:
        return 'Liberación Tarde'
    else:
        return 'Pendiente'

fcst['Liberacion_Status'] = fcst.apply(liberacion_status, axis=1)

# --- 8. Resumir para dashboard ---
resumen = (
    fcst.groupby(['ItemNo', 'Description', 'AC', 'ReqDate', 'ReleaseLimitDate', 'Liberacion_Status'])
    .size()
    .unstack(fill_value=0)
    .reset_index()
)

# --- 9. Exportar o visualizar ---
resumen.to_excel(r"C:\Users\j.vazquez\Desktop\PLAN_LIBERACION.xlsx", index=False)
print("Dashboard generado y guardado en Excel.")
