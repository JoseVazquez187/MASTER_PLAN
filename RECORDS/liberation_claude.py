#!/usr/bin/env python3
"""
Dashboard Simple del Plan de Liberación
Versión básica y funcional - CORREGIDA
"""

import flet as ft
import pandas as pd
import sqlite3
import numpy as np
from datetime import datetime, date, timedelta
import threading
import os

# ============================================================================
# CONFIGURACIÓN SIMPLE
# ============================================================================

DB_PATH = "J:/Departments/Operations/Shared/IT Administration/Python/IRPT/R4Database/R4Database.db"
DASHBOARD_DB = "simple_dashboard.db"
PORT = 8080

# Mapeo de entidades
ENTITY_MAPPING = {
    "EZ4301": "EZ4101", "EZ4303": "EZ4103", "EZ4304": "EZ4104",
    "EZ4305": "EZ4105", "EZ4306": "EZ4106", "EZ4307": "EZ4107",
    "EZ4309": "EZ4109", "EZ4310": "EZ4110", "EZ4311": "EZ4111",
    "EZ4312": "EZ4112", "EZ4313": "EZ4113", "EZ4314": "EZ4114",
    "EZ4316": "EZ4116", "EZ4317": "EZ4117", "EZ4318": "EZ4118",
    "EZ4324": "EZ4124", "EZ4325": "EZ4125"
}

# ============================================================================
# FUNCIONES SIMPLES
# ============================================================================

def cargar_datos():
    """Cargar datos básicos de la base de datos"""
    print("📊 Cargando datos...")
    print("🔌 Conectando a base de datos...")
    
    try:
        conn = sqlite3.connect(DB_PATH)
        print("✅ Conexión establecida")
        
        # Cargar forecast básico
        print("📥 Cargando forecast...")
        fcst_query = """
        SELECT Entity, AC, ItemNo, OpenQty, WO, ReqDate 
        FROM fcst 
        WHERE OpenQty > 0 
        LIMIT 1000
        """
        fcst = pd.read_sql_query(fcst_query, conn)
        print(f"✅ Forecast: {len(fcst)} registros cargados")
        
        # Cargar items básico
        print("📥 Cargando items...")
        items_query = """
        SELECT ItemNo, PLANTYPE, LT 
        FROM in92 
        WHERE ACTIVE = 'YES' 
        LIMIT 1000
        """
        items = pd.read_sql_query(items_query, conn)
        print(f"✅ Items: {len(items)} registros cargados")
        
        conn.close()
        print("🔐 Conexión cerrada")
        return fcst, items
        
    except Exception as e:
        print(f"❌ Error conectando BD: {e}")
        return None, None

def procesar_datos(fcst, items):
    """Procesar datos básicos"""
    print("🔄 Procesando datos...")
    print(f"📋 Merge de {len(fcst)} forecast con {len(items)} items...")
    
    # Merge básico
    data = fcst.merge(items, on='ItemNo', how='left')
    print(f"✅ Merge completado: {len(data)} registros")
    
    # Limpiar datos
    print("🧹 Limpiando datos...")
    data = data.fillna({'PLANTYPE': 'REF', 'LT': 5, 'WO': ''})
    
    # Crear llave
    print("🔑 Creando llaves...")
    data['llave'] = data['Entity'] + ' ' + data['AC'].astype(str)
    
    # Determinar liberación necesaria
    print("🔍 Analizando tipos de liberación...")
    liberation_needed = ['ASYFNL', 'ASYSTK', 'ASSY', 'ASYRT', 'ASYPH', 'ASYMS']
    data['necesita_liberacion'] = data['PLANTYPE'].isin(liberation_needed)
    
    # Status de liberación
    print("📊 Calculando status de liberación...")
    data['liberado'] = (data['WO'] != '') & (data['WO'].notna())
    
    # Resumen por AC
    print("📈 Generando resumen por AC...")
    resumen = data.groupby('llave').agg({
        'Entity': 'first',
        'AC': 'first',
        'ItemNo': 'count',
        'liberado': 'sum',
        'necesita_liberacion': 'sum'
    }).reset_index()
    
    resumen.columns = ['llave', 'entity', 'ac', 'total_items', 'items_liberados', 'items_necesitan_lib']
    
    # Calcular avance
    print("📐 Calculando avances...")
    resumen['avance'] = np.where(
        resumen['items_necesitan_lib'] > 0,
        resumen['items_liberados'] / resumen['items_necesitan_lib'],
        1.0
    )
    
    # Status final - MEJORADO
    print("🏁 Determinando status finales...")
    resumen['status'] = 'Pendiente'
    
    # Si no necesita liberación -> No Aplica
    resumen.loc[resumen['items_necesitan_lib'] == 0, 'status'] = 'No Aplica'
    
    # Si necesita liberación y está 100% liberado -> Liberado 100%
    resumen.loc[(resumen['items_necesitan_lib'] > 0) & (resumen['avance'] >= 1.0), 'status'] = 'Liberado 100%'
    
    # Si necesita liberación pero no está completo -> Pendiente
    resumen.loc[(resumen['items_necesitan_lib'] > 0) & (resumen['avance'] < 1.0), 'status'] = 'Pendiente'
    
    print(f"✅ Procesamiento completo: {len(resumen)} ACs procesados")
    return resumen

def guardar_en_db(resumen):
    """Guardar en base de datos simple"""
    try:
        conn = sqlite3.connect(DASHBOARD_DB)
        
        # Crear tabla simple
        conn.execute("""
            CREATE TABLE IF NOT EXISTS resumen_acs (
                llave TEXT PRIMARY KEY,
                entity TEXT,
                ac TEXT,
                total_items INTEGER,
                items_liberados INTEGER,
                items_necesitan_lib INTEGER,
                avance REAL,
                status TEXT,
                updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Limpiar y guardar
        conn.execute("DELETE FROM resumen_acs")
        resumen.to_sql('resumen_acs', conn, if_exists='append', index=False)
        
        conn.commit()
        conn.close()
        print("✅ Datos guardados en dashboard DB")
        return True
        
    except Exception as e:
        print(f"❌ Error guardando: {e}")
        return False

def cargar_desde_db():
    """Cargar datos desde dashboard DB"""
    try:
        conn = sqlite3.connect(DASHBOARD_DB)
        resumen = pd.read_sql_query("SELECT * FROM resumen_acs", conn)
        conn.close()
        return resumen
    except:
        return pd.DataFrame()

# ============================================================================
# APLICACIÓN FLET SIMPLE
# ============================================================================

class SimpleDashboard:
    def __init__(self):
        self.data = pd.DataFrame()
        self.data_filtrada = pd.DataFrame()
        self.loading = False
        
    def main(self, page: ft.Page):
        self.page = page
        page.title = "Dashboard Simple - Plan de Liberación"
        page.window_width = 1400
        page.window_height = 900
        page.padding = 20
        
        # Filtros
        self.filtro_entity = ft.Dropdown(
            label="Filtrar por Entity",
            width=150,
            options=[ft.dropdown.Option("Todos")],
            value="Todos",
            on_change=self.aplicar_filtros
        )
        
        self.filtro_status = ft.Dropdown(
            label="Filtrar por Status",
            width=150,
            options=[
                ft.dropdown.Option("Todos"),
                ft.dropdown.Option("Liberado 100%"),
                ft.dropdown.Option("Pendiente"),
                ft.dropdown.Option("No Aplica")
            ],
            value="Todos",
            on_change=self.aplicar_filtros
        )
        
        # Header con filtros
        header = ft.Row([
            ft.Text("🏭", size=30),
            ft.Text("Plan de Liberación - Dashboard Simple", size=24, weight=ft.FontWeight.BOLD),
            ft.Row([
                self.filtro_entity,
                self.filtro_status,
                ft.ElevatedButton(
                    "🔄 Actualizar",
                    on_click=self.actualizar_datos,
                    bgcolor="#2196F3",
                    color="#FFFFFF"
                )
            ], spacing=10)
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
        
        # Contenido principal
        self.content = ft.Column(spacing=20)
        
        page.add(header, ft.Divider(), self.content)
        
        # Cargar datos iniciales
        self.cargar_datos_iniciales()
    
    def cargar_datos_iniciales(self):
        """Cargar datos al inicio"""
        print("🚀 Iniciando carga de datos...")
        self.mostrar_loading()
        
        def cargar():
            try:
                print("🔍 Verificando datos locales...")
                # Intentar cargar desde DB local primero
                self.data = cargar_desde_db()
                
                if self.data.empty:
                    print("📥 No hay datos locales, cargando desde BD principal...")
                    # Si no hay datos, procesar desde DB principal
                    fcst, items = cargar_datos()
                    if fcst is not None and items is not None:
                        self.data = procesar_datos(fcst, items)
                        print("💾 Guardando en base local...")
                        guardar_en_db(self.data)
                        print("✅ Datos guardados exitosamente")
                    else:
                        print("❌ Error cargando datos de BD principal")
                        self.page.run_thread(lambda: self.mostrar_error("Error conectando a BD principal"))
                        return
                else:
                    print(f"✅ Datos locales encontrados: {len(self.data)} ACs")
                
                print("🖥️ Mostrando dashboard...")
                self.page.run_thread(lambda: [self.actualizar_filtros(), self.mostrar_dashboard()])
                
            except Exception as e:
                print(f"❌ Error en carga inicial: {e}")
                self.page.run_thread(lambda: self.mostrar_error(f"Error: {e}"))
        
        threading.Thread(target=cargar, daemon=True).start()
    
    def actualizar_filtros(self):
        """Actualizar opciones de filtros"""
        try:
            if not self.data.empty and hasattr(self, 'filtro_entity'):
                # Actualizar entities
                entities = sorted(self.data['entity'].unique())
                self.filtro_entity.options = [ft.dropdown.Option("Todos")] + [ft.dropdown.Option(e) for e in entities]
                
                # Aplicar filtros por defecto
                self.aplicar_filtros(None)
                print(f"✅ Filtros actualizados: {len(entities)} entities")
        except Exception as e:
            print(f"⚠️ Error actualizando filtros: {e}")
    
    def aplicar_filtros(self, e):
        """Aplicar filtros a los datos"""
        try:
            if self.data.empty:
                return
                
            # Copiar datos originales
            self.data_filtrada = self.data.copy()
            
            # Filtrar por entity
            if hasattr(self, 'filtro_entity') and self.filtro_entity.value and self.filtro_entity.value != "Todos":
                self.data_filtrada = self.data_filtrada[self.data_filtrada['entity'] == self.filtro_entity.value]
            
            # Filtrar por status
            if hasattr(self, 'filtro_status') and self.filtro_status.value and self.filtro_status.value != "Todos":
                self.data_filtrada = self.data_filtrada[self.data_filtrada['status'] == self.filtro_status.value]
            
            print(f"🔍 Filtros aplicados: {len(self.data_filtrada)} de {len(self.data)} ACs")
            
            # Actualizar dashboard si ya está cargado
            if hasattr(self, 'content') and len(self.content.controls) > 1:
                self.mostrar_dashboard()
        except Exception as e:
            print(f"⚠️ Error aplicando filtros: {e}")
    
    def actualizar_datos(self, e):
        """Actualizar datos manualmente"""
        print("🔄 Actualizando datos manualmente...")
        self.mostrar_loading()
        
        def actualizar():
            try:
                fcst, items = cargar_datos()
                if fcst is not None and items is not None:
                    self.data = procesar_datos(fcst, items)
                    guardar_en_db(self.data)
                    self.page.run_thread(lambda: [self.actualizar_filtros(), self.mostrar_dashboard()])
                else:
                    self.page.run_thread(lambda: self.mostrar_error("Error cargando datos"))
            except Exception as e:
                print(f"❌ Error en actualización: {e}")
                self.page.run_thread(lambda: self.mostrar_error(f"Error: {e}"))
        
        threading.Thread(target=actualizar, daemon=True).start()
    
    def mostrar_loading(self):
        """Mostrar indicador de carga con más detalle"""
        self.content.controls.clear()
        self.content.controls.append(
            ft.Container(
                content=ft.Column([
                    ft.ProgressRing(width=50, height=50),
                    ft.Text("⏳ Cargando datos...", size=18, weight=ft.FontWeight.BOLD),
                    ft.Text("Esto puede tomar unos segundos", size=14),
                    ft.Text("Revisa la consola para más detalles", size=12, color="#666666")
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=10),
                alignment=ft.alignment.center,
                height=300
            )
        )
        self.page.update()
    
    def mostrar_error(self, mensaje):
        """Mostrar error"""
        self.content.controls.clear()
        self.content.controls.append(
            ft.Container(
                content=ft.Column([
                    ft.Text("❌", size=50),  # Emoji en lugar de icono
                    ft.Text(mensaje, size=16, color="#F44336")  # Color directo
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                alignment=ft.alignment.center,
                height=200
            )
        )
        self.page.update()
    
    def mostrar_dashboard(self):
        """Mostrar dashboard principal"""
        datos_para_mostrar = self.data_filtrada if not self.data_filtrada.empty else self.data
        
        if datos_para_mostrar.empty:
            self.mostrar_error("No hay datos para mostrar")
            return
        
        self.content.controls.clear()
        
        # KPIs basados en datos filtrados
        total_acs = len(datos_para_mostrar)
        liberados = len(datos_para_mostrar[datos_para_mostrar['status'] == 'Liberado 100%'])
        pendientes = len(datos_para_mostrar[datos_para_mostrar['status'] == 'Pendiente'])
        no_aplica = len(datos_para_mostrar[datos_para_mostrar['status'] == 'No Aplica'])
        avance_promedio = datos_para_mostrar['avance'].mean() if len(datos_para_mostrar) > 0 else 0
        
        kpis = ft.Row([
            self.crear_kpi_card("Total ACs", str(total_acs), "📋", "#2196F3"),
            self.crear_kpi_card("Liberados", str(liberados), "✅", "#4CAF50"),
            self.crear_kpi_card("Pendientes", str(pendientes), "⏳", "#FF9800"),
            self.crear_kpi_card("No Aplica", str(no_aplica), "ℹ️", "#9E9E9E"),
            self.crear_kpi_card("Avance", f"{avance_promedio:.1%}", "📈", "#9C27B0")
        ], spacing=15)
        
        # Tabla con scroll
        tabla = self.crear_tabla_mejorada(datos_para_mostrar)
        
        self.content.controls.extend([
            ft.Text("Resumen Ejecutivo", size=20, weight=ft.FontWeight.BOLD),
            kpis,
            ft.Divider(),
            ft.Text(f"Detalle por AC - Mostrando {len(datos_para_mostrar)} ACs", size=18, weight=ft.FontWeight.BOLD),
            tabla
        ])
        
        self.page.update()
    
    def crear_kpi_card(self, titulo, valor, emoji, color):
        """Crear tarjeta KPI simple usando emojis"""
        return ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text(emoji, size=30),  # Emoji en lugar de icono
                    ft.Text(valor, size=20, weight=ft.FontWeight.BOLD, color=color),
                    ft.Text(titulo, size=12, color="#757575")  # Color directo
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
                padding=20,
                width=200,
                height=120
            ),
            elevation=2
        )
    
    def crear_tabla_mejorada(self, datos):
        """Crear tabla mejorada con scroll"""
        if datos.empty:
            return ft.Text("No hay datos")
        
        # Columnas
        columns = [
            ft.DataColumn(ft.Text("Entity", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("AC", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Total Items", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Liberados", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Necesita Lib", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Avance %", weight=ft.FontWeight.BOLD)),
            ft.DataColumn(ft.Text("Status", weight=ft.FontWeight.BOLD))
        ]
        
        # Todas las filas (no limitamos a 20)
        rows = []
        for _, row in datos.iterrows():
            # Color según status
            color = None
            if row['status'] == 'Liberado 100%':
                color = "#E8F5E8"  # Verde claro
            elif row['status'] == 'Pendiente':
                color = "#FFF3E0"  # Naranja claro
            elif row['status'] == 'No Aplica':
                color = "#F5F5F5"  # Gris claro
            
            # Verificar que las columnas existan
            entity_val = row.get('entity', row.get('Entity', 'N/A'))
            ac_val = row.get('ac', row.get('AC', 'N/A'))
            
            # Emoji para status
            status_emoji = ""
            if row['status'] == 'Liberado 100%':
                status_emoji = "✅ "
            elif row['status'] == 'Pendiente':
                status_emoji = "⏳ "
            elif row['status'] == 'No Aplica':
                status_emoji = "ℹ️ "
            
            rows.append(ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(str(entity_val), size=12)),
                    ft.DataCell(ft.Text(str(ac_val), size=12)),
                    ft.DataCell(ft.Text(str(row['total_items']), size=12)),
                    ft.DataCell(ft.Text(str(row['items_liberados']), size=12)),
                    ft.DataCell(ft.Text(str(row['items_necesitan_lib']), size=12)),
                    ft.DataCell(ft.Text(f"{row['avance']:.1%}", size=12, weight=ft.FontWeight.BOLD)),
                    ft.DataCell(ft.Text(f"{status_emoji}{row['status']}", size=12))
                ],
                color=color
            ))
        
        # Contenedor con scroll
        tabla_scroll = ft.Container(
            content=ft.DataTable(
                columns=columns,
                rows=rows,
                border=ft.border.all(1, "#E0E0E0"),
                border_radius=5,
                heading_row_color="#F5F5F5"
            ),
            height=400,
            width=1350,
            padding=10,
            border=ft.border.all(1, "#E0E0E0"),
            border_radius=10
        )
        
        return ft.Column([
            ft.Container(
                content=tabla_scroll,
                height=450,
                width=1350
            )
        ])

# ============================================================================
# EJECUTAR APLICACIÓN
# ============================================================================

def main():
    print("🚀 Iniciando Dashboard Simple...")
    
    # Verificar base de datos
    if not os.path.exists(DB_PATH):
        print(f"❌ Base de datos no encontrada: {DB_PATH}")
        print("Cambia la ruta en la variable DB_PATH del código")
        return
    
    print(f"✅ Base de datos encontrada")
    print(f"🌐 Abriendo en: http://localhost:{PORT}")
    
    # Ejecutar dashboard
    dashboard = SimpleDashboard()
    
    try:
        ft.app(target=dashboard.main, view=ft.AppView.FLET_APP, port=PORT)
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()

# ============================================================================
# INSTRUCCIONES SIMPLES
# ============================================================================
"""
INSTRUCCIONES:

1. INSTALAR:
   pip install flet pandas

2. EJECUTAR:
   python dashboard_simple.py

3. CAMBIAR RUTA SI ES NECESARIO:
   Edita la variable DB_PATH en la línea 15

4. USO:
   - Se abre automáticamente en el navegador
   - Botón "Actualizar" para refrescar datos
   - Muestra KPIs básicos y tabla de ACs
   - Colores: Verde = Liberado, Naranja = Pendiente

CAMBIOS REALIZADOS:
- ❌ Eliminados todos los ft.icons (causaban errores)
- ✅ Reemplazados con emojis
- ❌ Eliminadas referencias a ft.colors problemáticas
- ✅ Usados códigos de color directos (#RRGGBB)
- ✅ Compatibilidad mejorada con diferentes versiones de Flet

¡Simple y funcional sin errores! 🎉
"""