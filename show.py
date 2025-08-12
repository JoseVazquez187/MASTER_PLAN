import streamlit as st
from PIL import Image
import os
import requests
import plotly.graph_objects as go
import numpy as np
import matplotlib.pyplot as plt
# ===== Estilos ejecutivos y modo claro =====
st.markdown("""
    <style>
        html, body, [class*="css"] {
            background-color: #ffffff !important;
            color: #000000 !important;
            font-family: 'Segoe UI', sans-serif;
        }
        .stButton button {
            background-color: #e7e7e7;
            color: black;
            border-radius: 8px;
            border: 1px solid #ccc;
            padding: 0.5em 1em;
            font-size: 1em;
        }
        .stButton button:hover {
            background-color: #dcdcdc;
            cursor: pointer;
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
            padding-left: 4rem;
            padding-right: 4rem;
        }
        h1, h2, h3 {
            color: #1a1a1a;
        }
    </style>
""", unsafe_allow_html=True)

# ===== Control de navegación de slides =====
if "slide" not in st.session_state:
    st.session_state.slide = 1

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("⬅️ Anterior") and st.session_state.slide > 1:
        st.session_state.slide -= 1
with col2:
    if st.button("Siguiente ➡️"):
        st.session_state.slide += 1

slide = st.session_state.slide

# ===== Contenido por slide =====

if slide == 1:
    # Cargar logo
    logo = Image.open("images/LogoEZAirHD.JPG")
    
    # Crear columnas para logo + título
    col_logo, col_title = st.columns([1, 6])
    
    with col_logo:
        st.image(logo, width=60)
    with col_title:
        st.title("Inventory Reduction Plan Tools")
    st.subheader("Aplicaciones estrategicas para la reduccion del Inventario")
    st.markdown("""
        Reducir inventario **no es solo una meta financiera**.  
        Es una **estrategia de transformación operacional**.

        En tiempos de incertidumbre en la cadena de suministro, contar con las herramientas correctas  para **anticipar, decidir y accionar** es lo que diferencia a una empresa **eficiente** de una que solo **reacciona**.

        ---

        Con estas aplicaciones no solo buscamos eliminar exceso:  
        **Buscamos liberar capital, optimizar el flujo operativo y fortalecer la toma de decisiones.**

        ---

        <em style='font-size:18px'><strong>“Una herramienta que transforma datos en decisiones.
            No necesitas ser experto para tomar buenas decisiones: solo necesitas la herramienta correcta.” </strong></em>
        ---

        > Este conjunto de herramientas fue diseñado para aportar  
        > **visibilidad**, **velocidad de reacción**  y una **mejora continua en cómo planificamos, compramos y usamos material**.
        """, unsafe_allow_html=True)

elif slide == 2:
    import subprocess, os, sqlite3, pandas as pd, matplotlib.pyplot as plt
    from datetime import datetime

    st.title("🧠 Centro de Control")

    BASE_PATH = r"C:\Users\J.Vazquez\Desktop\Global_IRPT\MisScripts"

    projects = {
        "📦 Base de datos y expirados": {
            "📊 Actualizar Base de datos expirados": "db_expired.py",
            "📊 Get db Expirados": "get_db_expirados.py",
        },
        "🧱 Parche Expedite": {
            "🚚 Get DemandSource & shipdates Items_Boms": "get_demand_source_ItemBom.py",
            "🛠️ Correr Backschedule": "BSchedule_v3.py",
            "🔍 Parche Expedite": "parche_expedite.py",
            "📥 Parche to Original Expedite": "parche_to_original_expedite.py",
            "🔍 Parche Index actions": "expedite_index_actions.py",
            "🏭 Validacion parche expedite": "validacion_parche_expedite.py",
            "📥 Action PO'S": "Action_code_POs.py",
            "🏭 First Impact action POs": "Coverage_PO_action_first_Impact.py",
        },
        "📆 Planificación y Recertificación": {
            "📅 Plan de Recertificación": "recert_plan.py",
        },
        "📉 Loose Codes": {
            "🧮 Metrico de cancelaciones Aging PO'S": "cancelation_tool_aging.py",
            "🧮 MOQ Audit": "MOQ_Audit.py",
        }
    }

    def ejecutar_script(archivo):
        ruta_script = os.path.join(BASE_PATH, archivo)
        try:
            result = subprocess.run(["python", ruta_script], shell=True, capture_output=True, text=True, encoding='utf-8', errors='replace')
            if result.returncode == 0:
                st.success(f"✅ {archivo} se ejecutó correctamente.")
            else:
                st.error(f"❌ {archivo} terminó con errores.")
            if result.stdout:
                with st.expander("🖨️ Salida del script"):
                    st.code(result.stdout, language="bash")
            if result.stderr:
                with st.expander("⚠️ Errores del script"):
                    st.code(result.stderr, language="bash")
        except Exception as e:
            st.error(f"❌ Error al ejecutar {archivo}: {e}")

    def mostrar_link_documentacion(archivo_script):
        nombre_base = os.path.splitext(archivo_script)[0]
        for ext in [".pdf", ".md", ".txt"]:
            ruta_doc = os.path.join(BASE_PATH, nombre_base + ext)
            if os.path.exists(ruta_doc):
                with open(ruta_doc, "rb") as f:
                    st.download_button(
                        label="📄 Doc",
                        data=f,
                        file_name=os.path.basename(ruta_doc),
                        mime="application/octet-stream",
                        key=f"doc_{archivo_script}"
                    )
                return

    # === Pestañas internas ===
    tab_control, tab_reportes,carga_tab = st.tabs(["📋 Ejecutar Scripts", "📈 Reportes Descriptivos","📁 Carga de archivos"])

    # === TAB 1: Ejecutar scripts ===
    with tab_control:
        st.markdown("### 📂 Ejecutar Scripts por Categoría o Búsqueda")
        col1, col2 = st.columns([1, 5])
        with col1:
            search_term = st.text_input("🔍 Buscar script por nombre:")

        if search_term:
            found = False
            for categoria, scripts in projects.items():
                for nombre, archivo in scripts.items():
                    if search_term.lower() in nombre.lower():
                        found = True
                        col1, col2 = st.columns([5, 1])
                        with col1:
                            if st.button(nombre, key=f"run_{archivo}"):
                                ejecutar_script(archivo)
                        with col2:
                            mostrar_link_documentacion(archivo)
            if not found:
                st.warning("❌ No se encontraron scripts que coincidan.")
        else:
            for categoria, scripts in projects.items():
                with st.expander(categoria, expanded=False):
                    for nombre, archivo in scripts.items():
                        col1, col2 = st.columns([5, 1])
                        with col1:
                            if st.button(nombre, key=f"script_{archivo}"):
                                ejecutar_script(archivo)
                        with col2:
                            mostrar_link_documentacion(archivo)

    # === TAB 2: Reportes ejecutivos ===
    with tab_reportes:
        st.markdown("## 📈 Reportes Descriptivos")

        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            selected_date = st.date_input("📅 Ver datos hasta", value=datetime.today(), key="fecha_reporte")
        with col2:
            fig_width = st.number_input("📏 Ancho del gráfico", min_value=8, max_value=20, value=12, key="fig_width")
        with col3:
            fig_height = st.number_input("📐 Alto del gráfico", min_value=4, max_value=12, value=6, key="fig_height")

        if st.button("🔁 Actualizar Reportes", key="refresh_reportes"):
            try:
                db_path = os.path.join(BASE_PATH, "inventory.db")
                conn = sqlite3.connect(db_path)

                with st.expander("🧮 Resumen General de Expedite", expanded=True):
                    df = pd.read_sql("SELECT ItemNo, MLIKCode, Vendor FROM expedite", conn)
                    df['ItemNo'] = df['ItemNo'].astype(str).str.upper()
                    df['MLIKCode'] = df['MLIKCode'].astype(str).str.upper()
                    df['Vendor'] = df['Vendor'].astype(str).str.upper()
                    df_unique = df.drop_duplicates(subset='ItemNo')
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("🔢 Ítems únicos", df_unique['ItemNo'].nunique())
                    col2.metric("🅼 Total M", df_unique[df_unique['MLIKCode'] == 'M'].shape[0])
                    col3.metric("🅻 Total L", df_unique[df_unique['MLIKCode'] == 'L'].shape[0])
                    col4.metric("🏭 Proveedores únicos", df_unique['Vendor'].nunique())
                    st.dataframe(df_unique)

                # Aquí puedes agregar tus otros reportes: top 10 vendors, ABC, etc.

            except Exception as e:
                st.error(f"❌ Error al actualizar los reportes: {e}")
            finally:
                if 'conn' in locals():
                    conn.close()

    with carga_tab:
        st.title("📁 Carga de Archivos")
        if st.button("📥 Cargar Expedite", key="cargar_expedite"):
            ejecutar_script("cargar_expedite.py")
        if st.button("📥 Cargar Inventario IN 5 12", key="cargar_inventario"):
            ejecutar_script("cargar_in512.py")

elif slide == 3:
    st.title("📊 FEFO Tool Material con Shelf Life")
    st.write("Este slide muestra la informacion necesaria para el correcto funcionamiento de la Herrameinta.")
    st.info("💡 Clasifica, prioriza y programa de forma inteligente los materiales vencidos para su recertificación, evitar pérdida económica al destacar lo expirado y materiales en riesgo como no funcional.")
    

    with st.expander("📧 Ejemplo de correo de solicitud de validación"):
        st.image("images/foam.png", caption="Ejemplo de correo enviado para aprobación de SL%", use_container_width=True)

    st.markdown("### Librerias necesarias")
    with st.expander("📄 Instalar usando requirements.txt"):
        st.code("""
    # Paso 1: Crea un archivo llamado requirements.txt en la carpeta de tu proyecto.

    # Paso 2: Dentro del archivo, pega estas líneas:
    pandas
    numpy
    PyQt5
    matplotlib
    xlwings
    openpyxl

    # Paso 3: Abre tu terminal y ejecuta el siguiente comando para instalar todas las dependencias:
    pip install -r requirements.txt

    # Este comando le indica a pip que lea el archivo requirements.txt, instalando automáticamente todas las librerías listadas.

    # ✅ Así garantizas que el entorno esté listo para ejecutar la aplicación sin errores.
    """, language="python")

    st.markdown("### Reportes Necesarios")
    with st.expander("📄 Reportes necesarios"):
        st.markdown("""
        FEFO TOOL requiere los siguientes archivos para su correcta ejecución:

        1. <span style='color:#007acc'><b>cst54.txt</b></span>  
        Archivo de texto descargado de R4 donde muestran los costos de los materiales.  
        Se utiliza para priorizar por costo el plan de recertificación.

        2. <span style='color:#007acc'><b>in521.txt</b></span>  
        Archivo con información detallada de los lotes, incluyendo fechas de expiración.  
        Es esencial para aplicar la lógica FEFO (First Expired, First Out).

        3. <span style='color:#007acc'><b>Expirado.csv</b></span>  
        Reporte generado del SharePoint (Histórico cargado en SharePoint).  
        Se utiliza para descartar lotes que ya están previamente en SharePoint.

        4. <span style='color:#007acc'><b>historico_recertificacion.xlsx</b></span>  
        Registro de las recertificaciones ya realizadas.  
        Permite evitar reprocesos y entender la trazabilidad de decisiones pasadas.

        ---
        Todos estos archivos deben estar ubicados en la misma carpeta antes de ejecutar la herramienta:  
        <code>J:\\Departments\\Operations\\Shared\\IT Administration\\Python\\IRPT\\FEFO\\FEFO_files</code>
        """, unsafe_allow_html=True)


    st.markdown("### PDR Plan de Recertificacion")
    with st.expander("🎯 Plan de Recertificación (PDR)"):
        st.markdown("""
        <div style='font-family:Segoe UI, sans-serif; font-size:16px; line-height:1.6'>
        
        <span style='color:#1a73e8; font-weight:bold'>📌 Asignación automática de fechas de recertificación, basada en criterios clave:</span><br><br>

        🔸 <b style='color:#333'>Clasificación:</b> Todos los materiales vencidos o próximos a vencer son identificados y ordenados.<br>
        🔸 <b style='color:#d93025'>Vencidos:</b> Se priorizan de inmediato para evitar riesgos operativos.<br>
        🔸 <b style='color:#f9ab00'>Por vencer:</b> Se programan exactamente <b>3 días hábiles antes</b> de su fecha de expiración.<br>
        🔸 <b style='color:#188038'>Capacidad:</b> Se respeta la capacidad diaria de recertificación que el usuario define en la herramienta.<br><br>

        <i style='color:gray'>✅ El objetivo es anticiparse, reducir riesgo y evitar reprocesos innecesarios.</i>

        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        ### 🛠️ Lógica del Plan de Recertificación

        **1.- Primero se clasifica cada lote con base en su ubicación (`Bin`) o si ya está en SharePoint.**

        👉 Puedes ver las condiciones en la tabla mostrada a continuación.""")

        with st.expander("🧭 Tabla de clasificación por ubicación (Bin)"):
            st.image("images/pdr.png", caption="Condiciones lógicas para programar o excluir lotes", width=700)
        
        st.markdown("""**2.- Shelf Life = 0% → Prioridad crítica**

Separa los lotes “Programar” en dos grupos:

- 🟥 **Urgentes**: Shelf Life == 0 → ya vencidos → se agendan primero (respetando la capacidad diaria).
- 🟩 **Normales**: Shelf Life > 0 → se programan exactamente **3 días hábiles antes** de su fecha de vencimiento.

---

**3.- Orden por costo y agrupación**

• Esto asegura que **dentro del mismo Item y Lote, se dé prioridad a los lotes con mayor valor** (`Ext_cost`).

---

### 🎯 Criterios de Prioridad del Plan

- ✅ Solo se programan materiales **activos y viables**, excluyendo scrap, lab, RTV y SharePoint.
- 🔥 **Urgentes** (Shelf Life = 0%) se agendan primero, **respetando la capacidad diaria.**
- 📆 Lotes **por expirar** se programan 3 días antes de su vencimiento.
- 💰 Dentro de cada grupo, se prioriza el **mayor valor financiero (`Ext_cost`)**.""")

    st.markdown("### Como generar el PDR")
    with st.expander("📊 Login(PDR)"):
        st.markdown("""
### 🔐 Sistema de Acceso por Roles

La aplicación cuenta con un sistema de **inicio de sesión** que permite distinguir entre dos tipos de usuario:

- 👷‍♂️ **Producción**: acceso limitado solo a las funciones necesarias para ejecutar el plan de recertificación y consultar información relevante. No puede editar configuraciones ni modificar parámetros críticos.

- 🛠️ **Administrador**: acceso completo a todas las funciones, incluyendo ajustes de capacidad, filtros, validaciones, recálculos y reportes adicionales.

Este esquema permite mantener la operación segura y controlada, asegurando que solo personal autorizado pueda realizar cambios estructurales.
""")
        
        with st.expander("🖼️ Ver imagen de la pantalla de Login"):
            st.image("images/login.png", caption="pantalla login app tool", use_container_width=True)

        st.markdown("""### En el apartado de Admin""")
        st.markdown("""Antes de generar el PDR, asegúrate de actualizar los siguientes archivos y datos según se indica en esta tabla.""")
        
        with st.expander("🖼️ Actualizacion de reportes en el Path correspondiente"):
            st.image("images/panel_admin.png", caption="pantalla login app tool", use_container_width=True)

        with st.expander("🖼️ Panel Admin"):
            st.image("images/admin.png", caption="pantalla login app tool", use_container_width=True)


    st.markdown("### Plan de Recertificacion")
    with st.expander("📊 Archivo resultado"):
        

        ruta_excel = r"C:\Users\J.Vazquez\Desktop\show\ejemplo Excel\Plan de recertificacion.xlsx" 

        if st.button("📊 Abrir archivo Excel"):
            try:
                os.startfile(ruta_excel)  # Solo funciona en Windows
                st.success("✅ Archivo abierto correctamente.")
            except Exception as e:
                st.error(f"❌ Error al abrir el archivo: {e}")

elif slide == 4:
    st.title("🔁 Conversion de ING")
    st.write("Brindar al equipo de manufactura una visión estratégica de alternativas viables para la conversión de materiales obsoletos en materiales con demanda activa, facilitando decisiones informadas y orientadas a la optimización del inventario.")
    with st.expander("📧 Informacion necesaria"):
        st.markdown("""
    ### 📂 Archivos requeridos para la herramienta

    <HR style='border:1px solid #ccc'>

    🟦 <span style='font-size:17px'><b>Open Order</b></span><br>
    Contiene todas las órdenes de compra abiertas (**POs**) con cantidad pendiente.  
    Se utiliza para evaluar si la demanda está cubierta por órdenes en tránsito.<br>
    ➡️ <span style='color:#444'>Columna clave:</span> <b>OpnQ</b>

    <HR style='border:1px solid #ccc'>

    🟨 <span style='font-size:17px'><b>Slow Motion</b></span><br>
    Inventario actual por lote disponible en planta.  
    Este archivo permite ver cuales materiales son obsoletos o de lento movimiento.<br>
    ➡️ <span style='color:#444'>Columnas clave:</span> <b>Lot</b>, <b>Bin</b>, <b>QtyOH</b>

    <HR style='border:1px solid #ccc'>

    🟥 <span style='font-size:17px'><b>Expedite Report</b></span><br>
    Contiene los requerimientos de material por fecha de necesidad y demand source.  
    Usado como punto de partida para nuestro porceso de compra.<br>
    ➡️ <span style='color:#444'>Columnas clave:</span> <b>ItemNo</b>, <b>ReqQty</b>, <b>ReqDate</b>

    """, unsafe_allow_html=True)
    
    st.markdown("### Codigo de Python")
    
    with st.expander("📧Codigo python y logica de programacion"):
        st.markdown("## 🧠 Desglose completo de la lógica Python")

        st.markdown("""
        ### 🔹 Paso 1: Importar librerías y cargar datos principales
        Se ejecutan las funciones internas que extraen la base de datos de:
        - Pedidos abiertos (`Open Order`)
        - Inventario (`Slow Motion`)
        - Requerimientos urgentes (`Expedite`)
        """)

        st.code("""
        import pandas as pd
        import numpy as np
        import re
        import xlwings as xw

        oo = self.get_openOrder_information()
        slowmotion = self.get_items_information_from_slowmotion_table()
        expedite = self.get_expedite_information_indira()
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 2: Limpieza de datos e información de inventario
        - Se eliminan nulos y se convierte el inventario a números.
        - Se crea una columna que concatena `Bin`, `Lot` y `QtyOH` para visualización.
        """)

        st.code("""
        slowmotion = slowmotion.fillna('')
        slowmotion = slowmotion.loc[slowmotion['ItemNo'] != '']
        slowmotion['QtyOH'] = pd.to_numeric(slowmotion['QtyOH'], errors='coerce').fillna(0)
        slowmotion['ExtOH'] = pd.to_numeric(slowmotion['ExtOH'], errors='coerce').fillna(0)
        slowmotion['BinLotQty'] = "(" + slowmotion['Bin'].astype(str).str.strip() + ", " + slowmotion['Lot'].astype(str).str.strip() + ", " + slowmotion['QtyOH'].astype(float).astype(int).astype(str) + ")"
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 3: Agrupación de inventario por `ItemNo`
        """)

        st.code("""
        slowmotion_grouped = slowmotion.groupby('ItemNo').agg({
            'BinLotQty': lambda x: ', '.join(x),
            'Description': 'first',
            'ExtOH': 'sum',
            'QtyOH': 'sum',
            'MLI': 'first'
        }).reset_index()
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 4: Preparar Open Order por ítem
        """)

        st.code("""
        oo_PO = oo[['ItemNo','PONo','Ln','OpnQ']].copy()
        oo_PO['PONo'] = oo_PO['PONo'].astype(str)
        oo_PO['Ln'] = oo_PO['Ln'].astype(str)
        oo_PO['PO-ln'] = oo_PO['PONo'] + "-" + oo_PO['Ln']

        def agrupar_por_item(df):
            agrupado = df.groupby('ItemNo').agg({
                'PO-ln': lambda x: ','.join(x),
                'OpnQ': lambda x: ','.join(f"[{float(i):.2f}]" for i in x)
            }).reset_index()
            return agrupado.set_index('ItemNo')

        oo_PO = agrupar_por_item(oo_PO)
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 5: Pivotar Open Order total
        """)

        st.code("""
        oo = oo[['ItemNo', 'OpnQ']]
        oo['OpnQ'] = oo['OpnQ'].astype(float)
        pivot_table = oo.pivot_table(index='ItemNo', values='OpnQ', aggfunc='sum').reset_index()
        pivot_table = pivot_table.rename(columns={'OpnQ': 'Open_Order'}).replace('', 0)
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 6: Merge con Expedite y ordenar
        """)

        st.code("""
        base = pd.merge(expedite, pivot_table, how='left', on='ItemNo')
        base['ReqDate'] = pd.to_datetime(base['ReqDate'])
        base = base.sort_values(by=['ItemNo','ReqDate'])
        base = base[['EntityGroup', 'Project', 'AC','ItemNo','Description','PlanTp','OH','ReqQty','Open_Order','ReqDate','STDCost']]
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 7: Calcular Balance (solo OH y luego ajustado con PO)
        """)

        st.code("""
        base['Logic_item'] = 'y'
        base.loc[base['ItemNo'].duplicated(keep='first'), 'Logic_item'] = 'n'

        def balance_only_oh(df):
            df['Balance'] = 0.0
            now_current_balance = 0.0
            for index, row in df.iterrows():
                try:
                    oh = float(row['OH']) if str(row['OH']).strip() else 0.0
                    req = float(row['ReqQty']) if str(row['ReqQty']).strip() else 0.0
                except:
                    oh, req = 0.0, 0.0
                if row['Logic_item'] == 'y':
                    now_current_balance = oh - req
                else:
                    now_current_balance -= req
                df.at[index, 'Balance'] = now_current_balance
            return df

        base = balance_only_oh(base)
        """, language="python")

        st.code("""
        def adjust_balance_with_open_orders(df):
            df['Adjusted_Balance'] = df['Balance']
            for index, row in df.iterrows():
                try:
                    open_order = float(row['Open_Order'])
                except:
                    open_order = 0.0
                if row['Balance'] < 0 and open_order > 0:
                    df.at[index, 'Adjusted_Balance'] = row['Balance'] + open_order
            return df

        base = adjust_balance_with_open_orders(base)
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 8: Clasificación de cobertura
        """)

        st.code("""
        def status_by_coverage(df):
            conditions = [
                df['Balance'] >= 0,
                df['Adjusted_Balance'] >= 0
            ]
            choices = ['Coverage by OH', 'Coverage by PO']
            df['Status'] = np.select(conditions, choices, default='Shortage')
            return df

        base = status_by_coverage(base)
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 9: Agrupación final de estatus y truncamiento de claves
        """)

        st.code("""
        pivot_table = pd.pivot_table(
            base,
            index=['ItemNo','Description'],
            columns='Status',
            values='Adjusted_Balance',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        special_prefixes = ['AAA', 'AAB', 'AAC']

        def truncate_itemno(item):
            if not isinstance(item, str):
                return item
            if item.startswith("H") and item.count('-') == 1:
                return item.split('-')[0]
            for code in special_prefixes:
                index = item.find(code)
                if index != -1:
                    return item[:index]
            parts = item.split('-')
            if len(parts) >= 2:
                return '-'.join(parts[:2])
            return item
        """, language="python")

        st.code("""
        df_base = slowmotion_grouped.copy()
        df_expedite = pivot_table.copy()

        df_base['ItemNo_key'] = df_base['ItemNo'].apply(truncate_itemno)
        df_expedite['ItemNo_key'] = df_expedite['ItemNo'].apply(truncate_itemno)
        merged_df = pd.merge(df_base, df_expedite, on='ItemNo_key', how='left', suffixes=('_base', '_expedite'))
        """, language="python")

        st.markdown("""
        ### 🔹 Paso 10: Exportar resultados a Excel
        """)

        st.code("""
        def export_to_excel_safe(df, sheet):
            max_rows = 1048575
            if len(df) > max_rows:
                df = df.iloc[:max_rows]
            sheet.range('A1').value = df

        wb = xw.Book()
        s_base = wb.sheets.add(name="Base")
        export_to_excel_safe(slowmotion_grouped, s_base)
        s_exp = wb.sheets.add(name="expedite")
        export_to_excel_safe(pivot_table, s_exp)
        s_merge = wb.sheets.add(name="merged_df")
        export_to_excel_safe(merged_df, s_merge)
        """, language="python")

    st.markdown("### 📂 Abrir reporte en Excel")

    ruta_excel = r"C:\Users\J.Vazquez\Desktop\show\ejemplo Excel\Ing_change_alternatives.xlsx"  # <-- ajusta aquí la ruta

    if st.button("📊 Abrir archivo Excel"):
        try:
            os.startfile(ruta_excel)  # Solo funciona en Windows
            st.success("✅ Archivo abierto correctamente.")
        except Exception as e:
            st.error(f"❌ Error al abrir el archivo: {e}")

elif slide == 5:
    st.title("📦 Action Messages TOOL")
    st.markdown("""
### 🧩 Carga y Validación de Acciones de Cancelación

Esta herramienta facilita la identificación y trazabilidad de eventos clave relacionados con Action Codes 
                basados en el Aging de aparicion.

---

#### 🧠 ¿Qué hace esta aplicación?

- 📂 Recorre automáticamente una carpeta con archivos históricos (`Action-YYYYMMDD.xlsx`).
- 🔍 Filtra registros relevantes con base en los **códigos de acción seleccionados** (como `CN`, `AD`, `RI`...).
- 📅 Detecta la **primera aparición** de cada `PO_Linea` o `WorkOrder` con el código correspondiente.
- 🔗 Compara estos registros con el archivo actual de cancelaciones, calculando los **días transcurridos desde la primera acción**.
- 🗃️ Inserta y conserva la información procesada en una base de datos SQLite (`historico_actions.db`).

---

#### 🧾 Características adicionales

- 🎯 Menú interactivo para seleccionar los códigos a rastrear.
- 🚫 Evita duplicados al **omitir archivos ya procesados previamente**.
- 📊 Genera un reporte completo con el **aging por código de acción**.
- 🧱 Carga complementaria del archivo `pr 5 61.txt` para análisis extendido.

---

#### ✅ ¿Para qué sirve?

Esta herramienta es clave para:

- Detectar **retrasos operativos** y cuellos de botella mediante aging al no ejecutar el mensaje de accion.
- Generar **reportes trazables y auditables** con fechas respaldadas por evidencia documental.

> ⚙️ Ideal para equipos de **planeación y compras ** que requieren visibilidad sobre la ejecucion de los mensajes de accion.

""")

    with st.expander("🖼️ Pantalla principal"):
        st.image("images/cancelation_tool.png", caption="pantalla login app tool", use_container_width=True)

    with st.expander("🖼️Resultado en Base de datos "):
        st.image("images/cancelation_tool.png", caption="pantalla login app tool", use_container_width=True)

elif slide == 6:
    
    st.title("Material pendiente de Issue Plan de Junio por Sistema")
    st.markdown("### 📊 Clasificación de Cobertura de Kiteo")
    st.markdown("""
    | Clasificación         | Subcategoría                  | Líneas  | Monto ($)       |
    |----------------------|-------------------------------|---------|------------------|
    | ✅ **Kit Completo**   | **OH CUBRE**                  | 5,901   | **$413,912.37**  |
    |                      | OH después de 06/01/25        | 4,005   | $354,009.19      |
    |                      | VMI                           | 25      | $126.70          |
    | 🔶 **Incompleto**     | OH después de 06/01/25        | 8,927   | $1,026,505.15    |
    |                      | **OH CUBRE**                  | 10,296  | **$791,249.75**  |
    |                      | VMI                           | 28      | $517.10          |
    | 🔴 **No Cubierto**    |                               | 6,056   | **$1,720,213.37**|
    |                      |                               |         |                  |
    | **Total General**    |                               | 35,238  | **$4,306,533.63**|
    """)


    # Simulación de semanas (puedes ajustar según tus datos)
    # Semanas reales del mes de junio
    semanas = ['Semana 23', 'Semana 24', 'Semana 25', 'Semana 26', 'Semana 27']

    # Valores en millones (ajustados para visibilidad dentro del rango 29–35)
    inventario_real = [32.194, 32.45, 32.086, 32.67,32.62]
    inventario_ideal = [32.194, 31.7, 31.4, 31.1, 30.8]

    # Crear la gráfica
    fig, ax = plt.subplots()

    # Línea real
    ax.plot(semanas, inventario_real, label="Inventario Real", marker='o')
    for i, val in enumerate(inventario_real):
        ax.text(i, val + 0.2, f"{val:.1f}", ha='center', fontsize=9)

    # Línea ideal
    ax.plot(semanas, inventario_ideal, label="Inventario Ideal", linestyle='--', marker='x')
    for i, val in enumerate(inventario_ideal):
        ax.text(i, val - 0.4, f"{val:.1f}", ha='center', fontsize=9, color='gray')

    # Configuración del eje Y
    ax.set_ylim(30, 34)
    ax.set_title("📉 Evolución del Inventario – Junio")
    ax.set_ylabel("Inventario ($ millones)")
    #ax.set_xlabel("Semana")
    ax.grid(True)
    ax.legend()

    # Mostrar en Streamlit
    st.markdown("## 📊 Comparación: Inventario Real vs Inventario Ideal")
    st.pyplot(fig)

    st.markdown("#### Transformación del Surtido a Producción")
    
    
    
    col_intro1, col_intro2 = st.columns([2, 1])
    with col_intro1:
        st.markdown("""
        ### 🎯 Problemática Actual
        - Solicitudes manuales de materiales
        - Falta de priorización en el surtido
        - El material llega tarde a la línea
        - Pérdida de horas-hombre por tiempos de espera
        - Falta de visibilidad de órdenes completas (WO)
        - Surtido de ítems WO por WO (no consolidado)
        """)

    with col_intro2:
        st.markdown("#### 📦 Propuesta de Mejora")

        st.markdown("""
        - Solicitudes a través de la aplicación de Kiteo
        - Plan único de prioridades de producción
        - Seguimiento del avance de surtido (% surtido, nivel componente)
        - Reducción de horas-hombre perdidas por esperas
        - Material "Clear to Build" en línea de producción
        - Mayor eficiencia en la ejecución de los kiteos
        """)
    


    import streamlit as st

    st.markdown("### 🎯 Problema en el Surtido sin Priorización")

    st.markdown("""
    Cuando tenemos **un componente disponible** en inventario, pero **dos órdenes lo requieren**, el sistema actual lo muestra como disponible para ambas.  
    Esto provoca:

    - 🚨 **Confusión** al momento de kitar
    - ⌛ **Señales erroneas al momento de hacer la planeacion**
    - 📉 **Uso ineficiente del inventario**

    Ejemplo real:
    - Componente X está disponible solo **una vez** en almacén.
    - Dos órdenes (A y B) requieren ese componente.
    - El sistema marca ambas como listas para surtir.
    - En la realidad, **una queda incompleta**.
    """)

    # Diagrama del problema actual
    st.graphviz_chart('''
    digraph {
        node [shape=box style=filled fillcolor=lightgray]

        Inventario [label="Inventario: 1 pieza Componente X", fillcolor=orange]
        OrdenA [label="Orden A\nRequiere Componente X"]
        OrdenB [label="Orden B\nRequiere Componente X"]
        Kiteo [label="Kiteo recibe\nambas como listas", fillcolor=red]
        Resultado [label="Resultado: Solo una orden se puede surtir\nLa otra queda incompleta", fillcolor=pink]

        Inventario -> OrdenA
        Inventario -> OrdenB
        OrdenA -> Kiteo
        OrdenB -> Kiteo
        Kiteo -> Resultado
    }
    ''')


    st.markdown("""
    Implementar una herramienta que **aloque automáticamente el material disponible** según un plan único de prioridades permitiría:

    - ✅ Surtir solo la orden más prioritaria
    - 📈 Aumentar eficiencia y reducir reprocesos
    - 🔄 Mantener el inventario correctamente balanceado
    """)

    # Diagrama de solución
    st.graphviz_chart('''
    digraph {
        node [shape=box style=filled fillcolor=lightgray]

        Inventario [label="Inventario: 1 pieza Componente X", fillcolor=orange]
        Prioridad [label="Herramienta de Priorización\nOrden A es prioritaria", fillcolor=lightblue]
        OrdenA [label="Orden A\n(Surtir ahora)", fillcolor=lightgreen]
        OrdenB [label="Orden B\n(Esperar)", fillcolor=gray]
        Resultado [label="Resultado:\nSurtido correcto y controlado", fillcolor=lightyellow]

        Inventario -> Prioridad
        Prioridad -> OrdenA
        Prioridad -> OrdenB
        OrdenA -> Resultado
    }
    ''')

    flow_col2 = st.columns(1)[0]


    with flow_col2:
        st.markdown("##### ✅ Flujo Óptimo de Surtido ")

        st.graphviz_chart('''
    digraph {
        rankdir=TB;
        node [shape=ellipse, style=filled, fillcolor=white, fontname=Arial];

        // Flujo principal
        "Validación de Demanda Cargada" -> "Generación del Plan de Producción"
        "Generación del Plan de Producción" -> "Agrupación de Componentes por Orden"
        "Agrupación de Componentes por Orden" -> "Consolidación de Requerimientos"
        "Consolidación de Requerimientos" -> "Validación de Inventario Disponible"
        "Validación de Inventario Disponible" -> "Surtido por Prioridad (WO)"
        "Surtido por Prioridad (WO)" -> "Asignación de Carrito Digital"
        "Asignación de Carrito Digital" -> "Liberación para Almacén"
        "Liberación para Almacén" -> "Entrega Física a Línea"

        // Rama secundaria
        "Validación de Demanda Cargada" -> "MKL vs Cargado de FCST
        = Facturables fuera de MKL"

        // Posicionamiento horizontal
        { rank = same; "Validación de Demanda Cargada"; "MKL vs Cargado de FCST
        = Facturables fuera de MKL" }
    }
    ''')

    # Frase destacada
    st.markdown("""
    <div style="text-align:center; font-size:24px; font-weight:bold; margin-top:20px; color:#c0392b;">
        🔒 No hay producción sin visibilidad: <span style="text-decoration:underline;">si no lo veo, no lo produzco</span>.
    </div>
    """, unsafe_allow_html=True)

elif slide == 7:
    import streamlit as st
    import plotly.graph_objects as go

    st.set_page_config(page_title="Mater Plan/Kitting Plan", layout="centered")

    st.markdown("""
        <h1 style="text-align:center; color:#1a237e;">Progress Tracker</h1>

    """, unsafe_allow_html=True)

    phases = [
        "🛢️ Base de datos automatica",
        "🔮 Limpieza de Forecast",
        "📄 Limpieza de Sales Order",
        "🚪 Control de Entrada de Órdenes",
        "🔓 Control de Liberación de WO",
        "🛡️ Parche Expedite",
        "📅 BOM - Fecha de Necesidad / Ship Date",
        "🗂️ Backschedule por Nivel",
        "📝 Bitácoras de Entrega de material",
        "📲 Reporte de Escaneos",
        "🏭 Plan de Producción",
        "🔧 Plan de Producción Backshop",
        "⚙️ Plan de Producción Maquinado",
        "🛒 Requerimiento correcto de Comprado",
        "🎯 Lista de Prioridades por Proyecto",
        "📈 Generación de Cobertura de Materiales",
        "✅ Lista de Prioridades por Orden Clear",
        "🎯 Plan de Kiteo"
    ]

    # --- Fases que deseas que aparezcan ya seleccionadas (puedes modificar esta lista)
    initially_completed = [
        "🛢️ Base de datos automatica",
        "🗂️ Backschedule por Nivel",
        "📅 BOM - Fecha de Necesidad / Ship Date",
        "🛡️ Parche Expedite",
        
        # agrega aquí las que quieras que inicien marcadas
    ]

    st.markdown("### Select completed phases:")
    completed = []
    for phase in phases:
        checked = st.checkbox(phase, key=phase, value=(phase in initially_completed))
        completed.append(checked)

    num_completed = sum(completed)
    progress = int((num_completed / len(phases)) * 100)

    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = progress,
        gauge = {
            'axis': {'range': [0, 100]},
            'bar': {'color': "#2196f3"},
            'steps': [
                {'range': [0, 50], 'color': "#e3f2fd"},
                {'range': [50, 80], 'color': "#90caf9"},
                {'range': [80, 100], 'color': "#43a047"},
            ],
            'threshold': {
                'line': {'color': "red", 'width': 4},
                'thickness': 0.75,
                'value': 100
            }
        },
        title = {'text': "Overall Progress (%)"}
    ))
    fig.update_layout(height=350, margin=dict(t=20, b=20, l=20, r=20))

    st.plotly_chart(fig, use_container_width=True)
    st.progress(progress/100)
    st.markdown(f"""
    <div style="text-align:center;">
        <span style="font-size:2em; color:#1a237e;">{progress}% Completed</span>
    </div>
    """, unsafe_allow_html=True)

elif slide == 8:
    # ========== ESTILOS Y COLORES CORPORATIVOS ==========
    PRIMARY = "#1f2937"
    SECONDARY = "#4f8cff"
    ACCENT = "#fca311"
    BG = "#f4f7fa"
    GREEN = "#27ae60"  # Verde corporativo para resaltar pasos cumplidos o clave

    # === Aquí defines los pasos que deben ir en verde, puedes agregar más índices luego ===
    green_steps = [2,3,5,6,7,9]  # Por ejemplo: [2, 5, 7] para pasos 3, 6, 8

    st.set_page_config(
        page_title="MRP Revision para el calculo",
        layout="centered",
        page_icon="📊",
        initial_sidebar_state="collapsed"
    )

    st.markdown(f"""
        <style>
            .stApp {{ background-color: {BG}; }}
            .title-main {{
                font-size:2.3rem; font-weight:bold; color:{PRIMARY};
                margin-bottom:0.2em; letter-spacing:0.03em;
            }}
            .subtitle-main {{
                font-size:1.12rem; color:{SECONDARY}; margin-bottom:2em;
            }}
            .step-card {{
                background:white; border-radius:20px; box-shadow:0 2px 10px rgba(70,120,200,0.08);
                padding:20px 20px 12px 20px; margin-bottom:10px; min-height:90px;
                border-left:8px solid {SECONDARY};
            }}
            .step-title {{ font-size:1.12rem; font-weight:700; color:{PRIMARY}; }}
            .step-desc {{ color:#424242; font-size:0.99rem; margin-top:0.18em; }}
            .step-btn .element-container button {{
                width: 100% !important;
                border-radius: 14px !important;
                font-weight: bold;
                font-size: 1.05rem;
            }}
        </style>
    """, unsafe_allow_html=True)

    # ========== DATOS DE LOS PASOS ==========
    pasos = [
        ("01 FCST VS PP", "Todo lo que esté cargado en FCST debe estar en un plan de producción."),
        ("02 PLAN DE LIBERACIONES", "Las liberaciones de las WO deben estar conforme a necesidad de intermedios."),
        ("03 ALINEACIÓN DE FECHAS EN WO VS (FCST,SO)", "Toda WO liberada que su demanda sea FCST o SO debe estar alineada con la fecha de necesidad."),
        ("04 LIMPIEZA DE PSU", "La línea PSU requiere una limpieza manual, cada acción consta de 4 facturables, solo dos se cierran automáticamente."),
        ("05 LIMPIEZA DE P&S DE AC CON TERMINACIÓN 'B'", "El material P&S de los aviones con terminación B es necesario cerrarlos manualmente."),
        ("06 CREDIT MEMOS ABIERTOS", "Un credit memo abierto sin surtir es un potencial para duplicar demanda."),
        ("07 WO EN FIRME Y SIN EXPLOSIÓN DE MATERIALES", "Validación de materiales en WO Released."),
        ("08 REDUCCIÓN DE ÓRDENES A CANCELAR", "Eliminar órdenes a cancelar."),
        ("09 WO ACORDE A REVISIÓN EN R4", "Toda WO liberada debe de estar en la revisión más actual de R4."),
        ("10 NÚMEROS OBSOLETOS/EXPIRADOS EN LOCALIDAD NETEABLE", "Mover números obsoletos a localidades no neteables."),
        ("11 FORECAST SPARES PAST DUE", "Mover números obsoletos a localidades no neteables."),
    ]

    # ========== INICIALIZAR ESTADO DE NAVEGACIÓN ==========
    if "roadmap_selected" not in st.session_state:
        st.session_state["roadmap_selected"] = None

    def go_home():
        st.session_state["roadmap_selected"] = None

    # ========== LÓGICA DE NAVEGACIÓN ==========
    if st.session_state["roadmap_selected"] is None:
        # === VISTA GENERAL DEL ROADMAP ===
        st.markdown(
            '<div class="title-main">🔵 MRP Demand Checklist: 10 Pasos Críticos para la Excelencia Operativa</div>',
            unsafe_allow_html=True
        )
        st.markdown(
            '<div class="subtitle-main">Checklist: aseguremos que la planeación y ejecución de materiales sea impecable, visible y proactiva en cada etapa.</div>',
            unsafe_allow_html=True
        )
        st.markdown("""
                <div style="font-size:1.17rem; font-weight:bold; color:#1f2937; margin-bottom:10px;">
                    <div style="display:flex; align-items:center;">
                        <span style="background:#4f8cff; border-radius:6px; padding:5px 10px 5px 7px; margin-right:10px; font-size:1.08rem;">⚠️</span>
                        Validación <span style="color:#d35400; margin:0 4px 0 4px;">INDISPENSABLE</span>
                    </div>
                    <div style="margin-left:44px; font-weight:600; margin-top:2px;">
                        antes de correr el sistema o realizar cálculos.
                    </div>
                </div>
                """, unsafe_allow_html=True)

        cols = st.columns(2, gap="large")
        for idx, (titulo, desc) in enumerate(pasos):
            col = cols[0] if idx < 5 else cols[1]
            with col:
                with st.container():
                    # Si el paso está en green_steps, pintamos el título de verde y agregamos una paloma verde
                    if idx in green_steps:
                        st.markdown(f"""
                        <div class="step-card">
                            <div class="step-title" style="color:{GREEN};">{titulo} ✅</div>
                            <div class="step-desc">{desc}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div class="step-card">
                            <div class="step-title">{titulo}</div>
                            <div class="step-desc">{desc}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    if st.button("Ver detalle", key=f"btn_{idx}", use_container_width=True):
                        st.session_state["roadmap_selected"] = idx
        st.markdown("""
        ---
        <div style="text-align:center; color:#444; font-size:1.06rem; margin-top:2em;">
            <strong>¿Qué sigue?</strong> &nbsp;Evalúa en tiempo real cada paso de la checklist MRP para detectar riesgos y prioridades ejecutivas.<br>
            <span style="color:#4f8cff;">La excelencia en la planeación comienza con la visibilidad total.</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        # === VISTA DETALLE DE UN PASO ===
        idx = st.session_state["roadmap_selected"]
        titulo, desc = pasos[idx]
        # Si el paso está en green_steps, pintamos el título de verde y agregamos una paloma verde
        if idx in green_steps:
            st.markdown(
                f'<div class="title-main" style="color:{GREEN};">{titulo} ✅</div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                f'<div class="title-main">{titulo}</div>',
                unsafe_allow_html=True
            )
        st.markdown(
            f'<div class="subtitle-main">{desc}</div>',
            unsafe_allow_html=True
        )
        # Aquí puedes personalizar el contenido especial del detalle
        if idx == 5:
            with st.expander("📊 DASHBOARD", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\cm.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)
            st.markdown("""
            <div style="font-size:1.05rem; color:#4f8cff;">
                <b>¿Por qué es crítico?</b><br>
                Un Credit Memo abierto y sin surtir puede provocar duplicidad en la demanda y generar errores en la planeación de materiales. Es fundamental dar seguimiento puntual para evitar desbalance de inventarios y pedidos innecesarios.<br><br>
                <i>Valida semanalmente con el equipo de Finanzas y Compras los memos abiertos y confirma su cierre oportuno.</i>
            </div>
            """, unsafe_allow_html=True)
        
        

        if idx == 2:
            with st.expander("📊 DASHBOARD FORECAST", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\alineacionFCST.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)
            
            with st.expander("📊 DASHBOARD SALES ORDER", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\alineacionSalesOrder.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)   

            st.markdown("""
            <div style="font-size:1.05rem; color:#4f8cff;">
                <b>¿Por qué es crítico?</b><br>
                Despues de  liberada la WO la fecha de la WO (due-date) sera considerada como el Ship Date
            </div>
            """, unsafe_allow_html=True)
            with st.expander("📊 Demanda MA130 OHB FCST 8/2/2025 tabla de materiales 7/22/2025", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\demandadesalineada_ma130.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)   
                
            with st.expander("📊 Auditoria Masivo despues de la corrida", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\auditoria_cambio.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)   

        if idx == 6:
            with st.expander("📊 WO Sin Materiales", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\WOsinmateriales.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)
            st.markdown("""
            <div style="font-size:1.05rem; color:#4f8cff;">
                <b>¿Por qué es crítico?</b><br>
                Liberar una WO sin contar con materiales asignados equivale a generar inventario improductivo. El riesgo principal es identificar esta necesidad de material demasiado tarde, lo que provocará que el requerimiento entre al sistema de forma urgente y descontrolada.
            </div>
            """, unsafe_allow_html=True)

        if idx == 7:
            with st.expander("📊 WO a Cancelar Dashboard Principal", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\cancelationapp.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)
            
            with st.expander("📊 WO Analisis de cancelacion", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\analisisWOacancelar.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)

            st.markdown("""
            <div style="font-size:1.05rem; color:#4f8cff;">
                <b>¿Por qué es crítico?</b><br>
                Cancelar oportunamente las órdenes de trabajo innecesarias es fundamental para evitar el consumo indebido de materiales
                        , liberar capacidad operativa y prevenir distorsiones en la planeación. 
                        Si las órdenes a cancelar permanecen activas en el sistema, pueden generar 
                        requerimientos falsos de materiales, afectar los indicadores de inventario y 
                        provocar sobrecostos o retrasos innecesarios en otras prioridades del negocio.
            </div>
            """, unsafe_allow_html=True)


        if idx == 9:
            with st.expander("📊 Materiales Expirados en Localidad Neteable", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\materialexplocneteable.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)
            
            st.markdown("""
            <div style="font-size:1.05rem; color:#4f8cff;">
                <b>¿Por qué es crítico?</b><br>
                Distorsión en indicadores:
                Los inventarios neteables alimentan indicadores clave 
                        (como cobertura y disponibilidad). 
                        Tener materiales vencidos en estas Localidades da una imagen incorrecta 
                        de la salud del inventario.
            </div>
            """, unsafe_allow_html=True)

        if idx == 4:
            with st.expander("📊 Limpieza de Materiales con terminacion B", expanded=False):
                st.image(r"C:\\Users\\J.Vazquez\\Desktop\\show\\images\\limpiezainvoiced_B.png", caption="Evidencia de Credit Memos Abiertos", use_container_width=True)
            
            st.markdown("""
            <div style="font-size:1.05rem; color:#4f8cff;">
                <b>¿Por qué es crítico?</b><br>
                El sistema cierra el FCST solo si la cantidad facturada es igual a la cantidad cargada en el FCST 
                        El item,AC es igual a el que fue cargado en la PO
            </div>
            """, unsafe_allow_html=True)


        else:
            pass
        st.button("⬅️ Regresar al roadmap", on_click=go_home, use_container_width=True)





else:
    st.markdown("""## 👥 Responsables por Aplicación

A continuación se presentan los puntos de contacto responsables de cada herramienta o proceso dentro del Inventory Reduction Plan Tools:

| 🛠 Aplicación                              | 👤 Responsable         | 📧 Contacto                          |
|-------------------------------------------|------------------------|--------------------------------------|
| 📊 Base de Datos de Expirados             | José Vázquez           | [jose.vazquez@ezairinterior.com](mailto:jose.vazquez@ezairinterior.com) |
| 🚚 Parche Expedite                         | Laura Hernández        | [laura.hernandez@ezairinterior.com](mailto:laura.hernandez@ezairinterior.com) |
| 📅 Plan de Recertificación (PDR)          | Carlos Mendoza         | [carlos.mendoza@ezairinterior.com](mailto:carlos.mendoza@ezairinterior.com) |
| 🧮 Herramientas de Auditoría y KPI        | Mariana Ruiz           | [mariana.ruiz@ezairinterior.com](mailto:mariana.ruiz@ezairinterior.com) |
| 🤖 Modelos Predictivos               | Dojo team   | [inventory_dojo@ezairinterior.com](mailto:innovacion@ezairinterior.com) |

> 💡 En caso de requerir soporte o mejoras, contacta directamente con el responsable indicado.

---""")

    st.markdown("""
    ---
    ### 📫 ¿Tienes dudas o necesitas soporte adicional?

    Si tienes preguntas, comentarios o deseas colaborar con mejoras en estas herramientas, no dudes en ponerte en contacto:

    - 💼 **Contacto:** José Vázquez  
    - 📧 **Correo:** [jose.vazquez@tuempresa.com](mailto:jose.vazquez@tuempresa.com)  
    

    <div style='color:gray; font-size: 15px;'>
        “Estas herramientas están vivas: tu retroalimentación ayuda a mejorar continuamente.”
    </div>
    """, unsafe_allow_html=True)


    if st.button("🔄 Inicio"):
        st.session_state.slide = 1
