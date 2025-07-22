import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import os
import tkinter as tk
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
import threading

# === Configuración ===
# Detectar el nombre base del usuario
base_username = os.getlogin()

# # Crear los posibles paths
# username_with_one = base_username + ".ONE"
# user_path_with_one = os.path.join("C:\\Users", username_with_one)
# user_path_normal = os.path.join("C:\\Users", base_username)

# # Validar cuál usar
# if os.path.exists(user_path_with_one):
#     username = username_with_one
# else:
#     username = base_username

path_base = f'C:\\Users\\{base_username}\\Desktop\\PARCHE EXPEDITE'
file_fcst_excel = 'FCST.xlsx'
file_bom = 'SUPERBOM.txt'
file_lt = 'LeadTimes_in92.xlsx'
file_sort = 'SortCodes.xlsx'
file_holidays = 'Holidays.xlsx'
#checkpoint_file = os.path.join(path_base, "checkpoint.txt")

# === UI Principal ===
root = tk.Tk()
root.title("Progreso de Inserciones")
root.geometry("950x400")
log_output = ScrolledText(root, font=("Consolas", 10))
log_output.pack(expand=True, fill="both")

def log_insert(text):
    log_output.insert(tk.END, text + "\n")
    log_output.see(tk.END)
    root.update_idletasks()

# === Lógica de fechas hábiles ===
def restar_dias_habiles(fecha_inicio, dias, holiday_dates):
    contador = 0
    fecha_actual = fecha_inicio
    while contador < dias:
        fecha_actual -= timedelta(days=1)
        if fecha_actual.weekday() < 5 and fecha_actual.date() not in holiday_dates:
            contador += 1
    return fecha_actual

# === Función principal ===
def ejecutar_back_schedule():
    conn = None
    try:
        # === Leer y procesar FCST eliminando duplicados ===
        df_fcst = pd.read_excel(os.path.join(path_base, file_fcst_excel), usecols=["ItemNo", "ReqDate"])
        df_fcst['ItemNo'] = df_fcst['ItemNo'].astype(str).str.upper()
        df_fcst['ReqDate'] = pd.to_datetime(df_fcst['ReqDate'], errors='coerce')
        df_fcst.dropna(subset=['ItemNo', 'ReqDate'], inplace=True)

        # Quedarse solo con el ReqDate más lejano por ItemNo
        df_fcst = (
            df_fcst.sort_values('ReqDate', ascending=False)
                    .drop_duplicates(subset='ItemNo', keep='first')
                    .reset_index(drop=True)
        )

        # Calcular LT_TOTALIZADO respecto a una fecha ficticia
        fecha_ficticia = pd.Timestamp("2099-12-31")
        df_fcst['LT_TOTALIZADO'] = (fecha_ficticia - df_fcst['ReqDate']).dt.days

        # === Cargar BOM ===
        bom = pd.read_fwf(
            os.path.join(path_base, file_bom),
            widths=[23, 31, 20, 3, 2, 5, 3, 5, 10, 10, 13, 13, 11, 13, 10, 10, 11, 20],
            header=6
        ).fillna('')
        bom = bom.loc[bom['Level Number'] != 'End-of-Report.       1']

        # Asignar key por nivel 1
        bom['key'] = None
        ultima_key = None
        for idx, row in bom.iterrows():
            level_str = str(row['Level Number']).strip()
            try:
                level = int(level_str[0])
            except ValueError:
                continue
            if level == 1:
                ultima_key = row['Component']
            bom.at[idx, 'key'] = ultima_key

        bom['key'] = bom['key'].astype(str)
        bom['Sort'] = bom['Sort'].astype(str).str.upper()
        bom.reset_index(drop=True, inplace=True)
        bom['Orden_BOM_Original'] = bom.index + 1

        # === Reemplazo de componentes por excepciones ===
        path_exc = os.path.join(path_base, 'Exception_items.xlsx')
        if os.path.exists(path_exc):
            df_exc = pd.read_excel(path_exc)
            df_exc['Replace'] = df_exc['Replace'].astype(str).str.upper()
            df_exc['To'] = df_exc['To'].astype(str).str.upper()
            bom['Component'] = bom['Component'].astype(str).str.upper()
            bom = bom.replace({'Component': dict(zip(df_exc['Replace'], df_exc['To']))})

        # === Leer LT y Sort ===
        lt_df = pd.read_excel(os.path.join(path_base, file_lt))
        lt_df['ItemNo'] = lt_df['ItemNo'].astype(str).str.upper()
        lt_df = lt_df[['ItemNo','LeadTime']]

        lt_parche = pd.read_excel(os.path.join(path_base, 'LeadTime_Parche.xlsx'))
        lt_parche['ItemNo'] = lt_parche['ItemNo'].astype(str).str.upper()

        lt_df_final = pd.merge(lt_df, lt_parche, how='left', on='ItemNo')
        lt_df_final['LeadTime_Final'] = lt_df_final['LeadTime_y'].combine_first(lt_df_final['LeadTime_x'])
        lt_df_final = lt_df_final[['ItemNo', 'LeadTime_Final']].drop_duplicates('ItemNo')
        lt_df_final = lt_df_final.rename(columns={'LeadTime_Final': 'LT'})

        bom = pd.merge(bom, lt_df_final, how='left', left_on='Component', right_on='ItemNo')
        bom['LT'] = pd.to_numeric(bom['LT'], errors='coerce').fillna(0)

        sort_df = pd.read_excel(os.path.join(path_base, file_sort))
        sort_df = sort_df.rename(columns={sort_df.columns[0]: 'Sort', sort_df.columns[2]: 'Issue'})
        sort_df['Sort'] = sort_df['Sort'].astype(str).str.upper()

        bom = bom.merge(sort_df[['Sort', 'Issue']], how='left', on='Sort')

        # === Feriados ===
        holidays_df = pd.read_excel(os.path.join(path_base, file_holidays))
        holiday_dates = set(pd.to_datetime(holidays_df['Holidays']).dt.date)

        # === Crear base de datos ===
        db_path = os.path.join(path_base, f'Base1_{datetime.now().strftime("%Y_%m_%d_%H_%M")}.db')
        conn = sqlite3.connect(db_path)

        # === BackSchedule por ItemNo ===
        for i, row in df_fcst.iterrows():
            itemno = row['ItemNo']
            offset_dias = row['LT_TOTALIZADO']
            req_date = fecha_ficticia  # usar fecha ficticia como ShipDate
            log_insert(f"[{i+1}/{len(df_fcst)}] Procesando: {itemno} - OFFSET {offset_dias} dias")

            bom_item = bom[bom['key'] == itemno].copy()
            if bom_item.empty:
                log_insert(f"❌ Sin BOM: {itemno}")
                continue

            bom_item['ItemNo'] = itemno
            bom_item['ReqDate'] = req_date
            bom_item['ShipDate'] = req_date
            bom_item['BackScheduleDate'] = pd.NaT
            bom_item['Padre_M_usado'] = ""
            bom_item['LT_Padre_M'] = None

            for idx, row_b in bom_item.iterrows():
                nivel = int(str(row_b['Level Number']).strip()[0])
                mli = str(row_b.get('MLI', '')).strip().upper()
                lt = int(row_b['LT']) if not pd.isna(row_b['LT']) else 0

                if nivel == 1:
                    back_date = restar_dias_habiles(req_date, lt, holiday_dates) if lt > 0 else req_date
                    bom_item.at[idx, 'BackScheduleDate'] = back_date
                    bom_item.at[idx, 'Padre_M_usado'] = "Top Level"
                    bom_item.at[idx, 'LT_Padre_M'] = lt
                    continue

                padre_idx = None
                for j in range(idx - 1, -1, -1):
                    if (int(str(bom_item.at[j, 'Level Number']).strip()[0]) < nivel and 
                        str(bom_item.at[j, 'MLI']).strip().upper() == 'M' and 
                        pd.notna(bom_item.at[j, 'BackScheduleDate'])):
                        padre_idx = j
                        break

                if padre_idx is not None:
                    parent_fecha = bom_item.at[padre_idx, 'BackScheduleDate']
                    parent_lt = int(bom_item.at[padre_idx, 'LT']) if not pd.isna(bom_item.at[padre_idx, 'LT']) else 0
                    padre_nombre = bom_item.at[padre_idx, 'Component']
                else:
                    parent_fecha = req_date
                    parent_lt = 0
                    padre_nombre = "No encontrado"

                # === Ajuste para ítems L ===
                if mli == 'L':
                    if nivel == 2:
                        # L en nivel 2 => BS de 5 días
                        fecha_resultado = restar_dias_habiles(parent_fecha, 5, holiday_dates)
                    else:
                        # L en otro nivel => misma fecha del padre
                        fecha_resultado = parent_fecha
                elif mli == 'M':
                    # M siempre usa su propio LT
                    fecha_resultado = restar_dias_habiles(parent_fecha, lt, holiday_dates)
                else:
                    # Otros casos (por si hay vacíos u otros tipos), default: sin cambio
                    fecha_resultado = parent_fecha


                bom_item.at[idx, 'BackScheduleDate'] = fecha_resultado
                bom_item.at[idx, 'Padre_M_usado'] = padre_nombre
                bom_item.at[idx, 'LT_Padre_M'] = parent_lt




            bom_item['OFFSET_DIAS'] = (fecha_ficticia - bom_item['BackScheduleDate']).dt.days
            bom_item.to_sql('programacion_materiales', conn, index=False, if_exists='append')

    except Exception as e:
        log_insert(f"❌ Error: {str(e)}")
    finally:
        if conn:
            conn.close()
        log_insert("✅ Proceso completado.")
        root.after(1000, root.destroy)

# === Lanzar hilo ===
threading.Thread(target=ejecutar_back_schedule).start()
root.mainloop()
