import sqlite3
import pandas as pd
import os
from datetime import datetime, timedelta

def simple_wo_kiteo():
    print("🚀 SISTEMA DE PROPUESTA DE KITEO")
    print("=" * 50)
    db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join(desktop, f"Propuesta_Kiteo_{timestamp}.xlsx")

    try:
        print("📊 Conectando a la base de datos...")
        conn = sqlite3.connect(db_path)
        query = """
        SELECT 
            Entity,
            ItemNo,
            Description,
            WO,
            ReqDate,
            QtyFcst,
            OpenQty,
            Rev,
            UM,
            Proj
        FROM fcst 
        WHERE OpenQty > 0 
        AND WO IS NOT NULL 
        AND WO != ''
        ORDER BY ItemNo, Entity
        """

        df = pd.read_sql_query(query, conn)
        conn.close()
        print(f"✅ Datos cargados: {len(df)} registros")

        today = datetime.now().date()
        current_limit = today + timedelta(days=14)
        def simple_date_status(date_str):
            try:
                if pd.isna(date_str) or date_str == '' or date_str is None:
                    return "INVALID"
                date_obj = pd.to_datetime(date_str, errors='coerce').date()
                if pd.isna(date_obj):
                    return "INVALID"
                if date_obj < today:
                    return "PAST_DUE"
                elif date_obj <= current_limit:
                    return "CURRENT"
                else:
                    return "FUTURE"
            except:
                return "INVALID"
        df['DateStatus'] = df['ReqDate'].apply(simple_date_status)
        # === PROPUESTA FINAL: PAST_DUE separado y CURRENT+FUTURE juntos ===
        print("📦 Generando propuesta: PAST_DUE y CURRENT+FUTURE agrupados")

        # PAST_DUE
        df_pastdue = df[df['DateStatus'] == 'PAST_DUE'].copy()
        propuesta_pastdue = df_pastdue.groupby(['ItemNo', 'Entity', 'DateStatus']).agg({
            'WO': list,
            'Description': 'first',
            'OpenQty': 'sum',
            'QtyFcst': 'sum',
            'ReqDate': ['min', 'max'],
            'Rev': 'first',
            'UM': 'first',
            'Proj': 'first'
        }).reset_index()

        # CURRENT + FUTURE
        df_cf = df[df['DateStatus'].isin(['CURRENT', 'FUTURE'])].copy()
        df_cf['DateStatus'] = 'CURRENT+FUTURE'
        propuesta_cf = df_cf.groupby(['ItemNo', 'Entity', 'DateStatus']).agg({
            'WO': list,
            'Description': 'first',
            'OpenQty': 'sum',
            'QtyFcst': 'sum',
            'ReqDate': ['min', 'max'],
            'Rev': 'first',
            'UM': 'first',
            'Proj': 'first'
        }).reset_index()

        # Combinar ambos bloques
        propuesta_final = pd.concat([propuesta_pastdue, propuesta_cf], ignore_index=True)
        propuesta_final.columns = ['ItemNo', 'Entity', 'DateStatus', 'WO_List', 'Description',
                                   'OpenQty_Total', 'QtyFcst_Total', 'ReqDate_Min', 'ReqDate_Max',
                                   'Rev', 'UM', 'Proj']

        propuesta_final['WO_Count'] = propuesta_final['WO_List'].apply(len)
        propuesta_final['WO_String'] = propuesta_final['WO_List'].apply(lambda x: ', '.join(map(str, x)))
        propuesta_final = propuesta_final.reset_index(drop=True)
        propuesta_final['Group_ID'] = range(1, len(propuesta_final) + 1)

        # Ordenar primero PAST_DUE luego CURRENT+FUTURE
        propuesta_final['DateStatus_Order'] = propuesta_final['DateStatus'].map(
            {'PAST_DUE': 1, 'CURRENT+FUTURE': 2}
        )
        propuesta_final = propuesta_final.sort_values(by=['DateStatus_Order', 'ItemNo']).drop(columns='DateStatus_Order')

        # === GUARDAR EN EXCEL SOLO PROPUESTA_KITEO ===
        print("💾 Guardando archivo Excel...")

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            propuesta_cols = ['Group_ID', 'ItemNo', 'Description', 'Entity', 'WO_Count',
                              'WO_String', 'OpenQty_Total', 'QtyFcst_Total', 'DateStatus',
                              'ReqDate_Min', 'ReqDate_Max', 'Rev', 'UM']
            propuesta_final[propuesta_cols].to_excel(writer, sheet_name='Propuesta_Kiteo', index=False)

        print(f"\n🎉 PROCESO COMPLETADO EXITOSAMENTE")
        print(f"📁 Archivo guardado: {output_file}")

        try:
            os.startfile(output_file)
        except:
            print("⚠️ Abrir manualmente el archivo desde el escritorio")

        return True

    except Exception as e:
        print(f"❌ ERROR: {e}")
        return False

if __name__ == "__main__":
    simple_wo_kiteo()
    input("\nPresiona Enter para salir...")
