import flet as ft
import sqlite3
import pandas as pd
import datetime
import threading
import time
from typing import Dict, List, Optional, Tuple

class DatabaseConfig:
    """Configuración escalable de la base de datos"""
    def __init__(self):
        self.db_path = r"C:\Users\J.Vazquez\Desktop\R4Database.db"
        self.connection_timeout = 30
        self.auto_refresh = True
        self.refresh_interval = 300
        self.date_range_days = 90

class DatabaseManager:
    """Gestor de conexión y consultas a la base de datos"""
    def __init__(self, config: DatabaseConfig):
        self.config = config
        self.last_update = None
        self.connection_status = "disconnected"
        self.po_timeline_data = pd.DataFrame()
        
    def get_main_query(self):
        """Query principal para obtener datos del Line of Balance incluyendo Purchase Orders"""
        return """
        SELECT
            expedite.EntityGroup,
            expedite.Project,
            expedite.AC,
            expedite.ItemNo,
            expedite.Description,
            expedite.PlanTp,
            expedite.ReqQty,
            expedite.Vendor,
            expedite.ReqDate,
            expedite.ShipDate,
            expedite.MLIKCode,
            expedite.LT,
            expedite."Std-Cost",
            expedite.LotSize,
            expedite.UOM,
            entity.EntityName,
            IFNULL(oh_summary.Total_OH, 0) AS OH_Net,
            IFNULL(rwk_summary.RWK_OH, 0) AS OH_RWK,
            IFNULL(qa_summary.QA_OH, 0) AS OH_QA,
            (IFNULL(oh_summary.Total_OH, 0) + IFNULL(rwk_summary.RWK_OH, 0) + IFNULL(qa_summary.QA_OH, 0)) AS OH_Total,
            IFNULL(po_summary.Total_PO_Qty, 0) AS PO_Qty,
            IFNULL(po_summary.PO_Count, 0) AS PO_Count,
            po_summary.Next_PO_Date,
            po_summary.PO_Details,
            'W' || strftime('%W', expedite.ReqDate) || '-' || strftime('%Y', expedite.ReqDate) AS WeekYear
        FROM expedite
        LEFT JOIN entity ON expedite.EntityGroup = entity.EntityGroup
        LEFT JOIN (
            SELECT ItemNo, SUM(OH) AS Total_OH
            FROM oh_pr_table
            GROUP BY ItemNo
        ) AS oh_summary ON expedite.ItemNo = oh_summary.ItemNo
        LEFT JOIN (
            SELECT ItemNo, SUM(OH) AS RWK_OH
            FROM ReworkLoc_all
            GROUP BY ItemNo
        ) AS rwk_summary ON expedite.ItemNo = rwk_summary.ItemNo
        LEFT JOIN (
            SELECT ItemNo, SUM(OH) AS QA_OH
            FROM qa_oh
            GROUP BY ItemNo
        ) AS qa_summary ON expedite.ItemNo = qa_summary.ItemNo
        LEFT JOIN (
            SELECT 
                ItemNo,
                SUM(OpnQ) AS Total_PO_Qty,
                COUNT(*) AS PO_Count,
                MIN(RevPrDt) AS Next_PO_Date,
                GROUP_CONCAT(
                    'PO:' || PONo || 
                    '|Qty:' || OpnQ || 
                    '|Date:' || RevPrDt || 
                    '|Vendor:' || IFNULL(Vendor, 'N/A'), 
                    ';'
                ) AS PO_Details
            FROM openOrder
            WHERE OpnQ > 0 
            AND RevPrDt IS NOT NULL
            AND RevPrDt >= date('now')
            GROUP BY ItemNo
        ) AS po_summary ON expedite.ItemNo = po_summary.ItemNo
        WHERE expedite.MLIKCode = "L"
        AND expedite.PlanTp <> "VMIHDW"
        ORDER BY expedite.AC, expedite.Project, expedite.ItemNo
        """
    
    def get_po_timeline_query(self):
        """Query específico para obtener timeline detallado de POs"""
        date_range = self.config.date_range_days  # Usar el atributo correcto
        return f"""
        SELECT 
            PONo,
            ItemNo,
            OpnQ,
            RevPrDt,
            Vendor,
            BuyerName,
            LineMemo,
            NoteStatus,
            StdCost,
            (OpnQ * StdCost) AS PO_Value
        FROM openOrder
        WHERE OpnQ > 0
        AND RevPrDt IS NOT NULL
        AND RevPrDt >= date('now', '-{date_range} days')
        AND RevPrDt <= date('now', '+{date_range} days')
        ORDER BY ItemNo, RevPrDt
        """
    
    def connect_and_fetch_data(self):
        """Conectar a la base de datos y obtener los datos"""
        try:
            print("🔄 Cambiando estado a 'connecting'...")
            self.connection_status = "connecting"
            
            print(f"🔄 Conectando a SQLite: {self.config.db_path}")
            conn = sqlite3.connect(self.config.db_path, timeout=self.config.connection_timeout)
            
            print("🔄 Ejecutando query principal...")
            main_df = pd.read_sql_query(self.get_main_query(), conn)
            print(f"✅ Query principal completado: {len(main_df)} registros")
            
            print("🔄 Ejecutando query de POs...")
            po_timeline_df = pd.read_sql_query(self.get_po_timeline_query(), conn)
            print(f"✅ Query PO completado: {len(po_timeline_df)} registros")
            
            print("🔄 Procesando datos...")
            main_df = self.process_data(main_df)
            po_timeline_df = self.process_po_data(po_timeline_df)
            
            conn.close()
            print("✅ Conexión cerrada")
            
            self.connection_status = "connected"
            self.last_update = datetime.datetime.now()
            self.po_timeline_data = po_timeline_df
            
            print("✅ Datos procesados correctamente")
            return main_df, po_timeline_df
            
        except Exception as e:
            self.connection_status = "error"
            print(f"❌ Error en connect_and_fetch_data: {str(e)}")
            raise Exception(f"Error conectando a la base de datos: {str(e)}")
    
    def process_po_data(self, po_df):
        """Procesar datos de Purchase Orders"""
        if po_df.empty:
            return po_df
        
        po_df['RevPrDt'] = pd.to_datetime(po_df['RevPrDt'], errors='coerce')
        po_df['Days_to_Arrival'] = (po_df['RevPrDt'] - datetime.datetime.now()).dt.days
        
        def get_po_status(days):
            if days < 0:
                return 'overdue'
            elif days <= 7:
                return 'critical'
            else:
                return 'normal'
        
        po_df['PO_Status'] = po_df['Days_to_Arrival'].apply(get_po_status)
        return po_df
    
    def process_data(self, df):
        """Procesar datos para cálculos de Line of Balance"""
        if df.empty:
            return df
        
        df['ReqDate'] = pd.to_datetime(df['ReqDate'], errors='coerce')
        df['ShipDate'] = pd.to_datetime(df['ShipDate'], errors='coerce')
        df['Next_PO_Date'] = pd.to_datetime(df['Next_PO_Date'], errors='coerce')
        
        df['DailyDemand'] = df['ReqQty'] / 30
        
        def calc_current_coverage(row):
            if row['DailyDemand'] > 0:
                return row['OH_Total'] / row['DailyDemand']
            return 0
        
        df['Current_Coverage'] = df.apply(calc_current_coverage, axis=1)
        
        def calc_coverage_with_po(row):
            if row['DailyDemand'] <= 0:
                return 0
            
            current_stock = row['OH_Total']
            daily_demand = row['DailyDemand']
            po_qty = row['PO_Qty'] if pd.notna(row['PO_Qty']) else 0
            
            if po_qty == 0:
                return current_stock / daily_demand
            
            days_current_stock = current_stock / daily_demand if daily_demand > 0 else 0
            
            if pd.notna(row['Next_PO_Date']):
                days_to_po = (row['Next_PO_Date'] - datetime.datetime.now()).days
            else:
                days_to_po = 999
            
            if days_to_po <= days_current_stock:
                return (current_stock + po_qty) / daily_demand
            else:
                return current_stock / daily_demand
        
        df['Coverage_with_PO'] = df.apply(calc_coverage_with_po, axis=1)
        df['Coverage'] = df['Coverage_with_PO']
        
        def get_status(coverage):
            if coverage < 5:
                return 'critical'
            elif coverage < 15:
                return 'warning'
            else:
                return 'safe'
        
        df['Status'] = df['Coverage'].apply(get_status)
        
        def calc_days_to_po(row):
            if pd.notna(row['Next_PO_Date']):
                return (row['Next_PO_Date'] - datetime.datetime.now()).days
            return None
        
        df['Days_to_Next_PO'] = df.apply(calc_days_to_po, axis=1)
        df['InventoryValue'] = df['OH_Total'] * df['Std-Cost']
        df['PO_Value'] = df['PO_Qty'] * df['Std-Cost']
        df['IsActivePO'] = df['PO_Qty'] > 0
        df['Has_Overdue_PO'] = df['Days_to_Next_PO'] < 0
        
        return df

class LOBAnalyzer:
    """Analizador de datos para Line of Balance"""
    def __init__(self, data):
        self.data = data
    
    def get_kpis(self, buyer_filter="Todos", project_filter="Todos"):
        """Calcular KPIs principales"""
        filtered_data = self.filter_data(buyer_filter, project_filter)
        
        if filtered_data.empty:
            return {
                'total_items': 0,
                'critical_items': 0,
                'avg_coverage': 0,
                'total_inventory_value': 0,
                'active_pos': 0
            }
        
        return {
            'total_items': len(filtered_data),
            'critical_items': len(filtered_data[filtered_data['Status'] == 'critical']),
            'avg_coverage': round(filtered_data['Coverage'].mean(), 1),
            'total_inventory_value': filtered_data['InventoryValue'].sum(),
            'active_pos': len(filtered_data[filtered_data['IsActivePO']])
        }
    
    def filter_data(self, buyer_filter="Todos", project_filter="Todos"):
        """Filtrar datos según criterios"""
        filtered = self.data.copy()
        
        if buyer_filter != "Todos":
            filtered = filtered[filtered['AC'] == buyer_filter]
        
        if project_filter != "Todos":
            filtered = filtered[filtered['Project'] == project_filter]
        
        return filtered
    
    def get_critical_items(self, buyer_filter="Todos", project_filter="Todos", limit=20):
        """Obtener items críticos ordenados por cobertura"""
        filtered_data = self.filter_data(buyer_filter, project_filter)
        
        critical_data = filtered_data[
            (filtered_data['Status'] == 'critical') | 
            (filtered_data['Status'] == 'warning')
        ].sort_values('Coverage')
        
        return critical_data.head(limit)

class SimpleChartManager:
    """Gestor de gráficos simples sin Plotly"""
    def __init__(self):
        pass
    
    def create_simple_chart(self, data, period_days=90):
        """Crear gráfico simple de cobertura usando colores del tema"""
        if data.empty:
            return ft.Text("No hay datos disponibles", text_align=ft.TextAlign.CENTER)
        
        timeline_data = self.generate_simple_timeline(data, period_days)
        chart_bars = []
        max_coverage = max([item['coverage'] for item in timeline_data]) if timeline_data else 100
        
        for i, item in enumerate(timeline_data[:30]):
            height = (item['coverage'] / max_coverage) * 200 if max_coverage > 0 else 0
            
            if item['coverage'] < 5:
                color = "#FF0000"
            elif item['coverage'] < 15:
                color = "#FFA500"
            else:
                color = "#00FF00"
            
            bar = ft.Container(
                content=ft.Text(f"{item['coverage']:.0f}", size=8, text_align=ft.TextAlign.CENTER),
                width=15,
                height=max(height, 10),
                bgcolor=color,
                border_radius=2,
                margin=1,
                tooltip=f"Día {i+1}: {item['coverage']:.1f} días de cobertura"
            )
            chart_bars.append(bar)
        
        return ft.Container(
            content=ft.Column([
                ft.Text("Proyección de Cobertura (30 días)", weight=ft.FontWeight.BOLD),
                ft.Row([
                    ft.Text("Crítico (<5d)", color="#FF0000", size=10),
                    ft.Text("Warning (<15d)", color="#FFA500", size=10),
                    ft.Text("Seguro (>15d)", color="#00FF00", size=10),
                ], spacing=10),
                ft.Container(
                    content=ft.Row(chart_bars, spacing=2, scroll=ft.ScrollMode.AUTO),
                    height=220,
                    border=ft.border.all(1, "#CCCCCC"),
                    border_radius=5,
                    padding=10
                )
            ], spacing=10),
            padding=10
        )
    
    def generate_simple_timeline(self, data, period_days=90):
        """Generar timeline simplificado"""
        timeline = []
        
        if data.empty:
            return timeline
        
        total_inventory = data['OH_Total'].sum()
        total_daily_demand = data['DailyDemand'].sum()
        
        for i in range(period_days):
            remaining_inventory = max(0, total_inventory - (total_daily_demand * i))
            coverage = remaining_inventory / total_daily_demand if total_daily_demand > 0 else 0
            
            timeline.append({
                'day': i + 1,
                'coverage': coverage,
                'inventory': remaining_inventory
            })
        
        return timeline

class LineOfBalanceApp:
    """Aplicación principal de Line of Balance"""
    def __init__(self):
        self.config = DatabaseConfig()
        self.db_manager = DatabaseManager(self.config)
        self.data = pd.DataFrame()
        self.po_data = pd.DataFrame()
        self.analyzer = None
        self.chart_manager = SimpleChartManager()
        self.selected_buyer = "Todos"
        self.selected_project = "Todos"
        
        # Estado del tema
        self.is_dark_mode = False
        
    def get_theme_colors(self):
        """Obtener colores según el tema actual - Paleta ejecutiva"""
        if self.is_dark_mode:
            return {
                'bg_primary': "#0f172a",       # Slate-900
                'bg_secondary': "#1e293b",     # Slate-800
                'bg_card': "#334155",          # Slate-700
                'text_primary': "#f8fafc",     # Slate-50
                'text_secondary': "#cbd5e1",   # Slate-300
                'border': "#475569",           # Slate-600
                'accent': "#3b82f6",           # Blue-500
            }
        else:
            return {
                'bg_primary': "#f8fafc",       # Slate-50
                'bg_secondary': "#ffffff",     # White
                'bg_card': "#ffffff",          # White
                'text_primary': "#0f172a",     # Slate-900
                'text_secondary': "#64748b",   # Slate-500
                'border': "#e2e8f0",           # Slate-200
                'accent': "#3b82f6",           # Blue-500
            }
    
    def toggle_theme(self, e):
        """Alternar entre tema claro y oscuro"""
        self.is_dark_mode = not self.is_dark_mode
        
        # Actualizar tema de la página
        if self.is_dark_mode:
            self.page.theme_mode = ft.ThemeMode.DARK
            self.theme_toggle_button.icon = "light_mode"
            self.theme_toggle_button.tooltip = "Cambiar a modo claro"
        else:
            self.page.theme_mode = ft.ThemeMode.LIGHT
            self.theme_toggle_button.icon = "dark_mode"
            self.theme_toggle_button.tooltip = "Cambiar a modo oscuro"
        
        # Recrear la interfaz con los nuevos colores
        self.rebuild_ui()
        self.page.update()
        
    def main(self, page: ft.Page):
        """Función principal de la aplicación Flet"""
        page.title = "Line of Balance Dashboard - R4Database"
        page.theme_mode = ft.ThemeMode.LIGHT
        page.padding = 20
        page.scroll = ft.ScrollMode.AUTO
        
        self.page = page
        
        # Variables de estado de la UI
        self.connection_status_text = ft.Text("Desconectado", color="#FF0000")
        self.last_update_text = ft.Text("Nunca")
        
        # Botón de toggle de tema
        self.theme_toggle_button = ft.IconButton(
            icon="dark_mode",
            tooltip="Cambiar a modo oscuro",
            on_click=self.toggle_theme
        )
        
        # Filtros
        self.buyer_dropdown = ft.Dropdown(
            label="Comprador",
            options=[ft.dropdown.Option("Todos")],
            value="Todos",
            on_change=self.on_filter_change,
            width=200
        )
        
        self.project_dropdown = ft.Dropdown(
            label="Proyecto",
            options=[ft.dropdown.Option("Todos")],
            value="Todos",
            on_change=self.on_filter_change,
            width=200
        )
        
        # Filtro de rango de fechas
        self.date_range_dropdown = ft.Dropdown(
            label="Rango de POs",
            options=[
                ft.dropdown.Option("30", "30 días"),
                ft.dropdown.Option("60", "60 días"),
                ft.dropdown.Option("90", "90 días"),
                ft.dropdown.Option("120", "120 días"),
                ft.dropdown.Option("180", "6 meses"),
                ft.dropdown.Option("365", "1 año")
            ],
            value="90",
            on_change=self.on_date_range_change,
            width=150
        )
        
        # Inicializar la UI
        self.build_initial_ui()
        
        # Cargar datos inicial
        self.refresh_data()
    
    def build_initial_ui(self):
        """Construir la interfaz inicial"""
        self.kpi_cards = self.create_kpi_cards()
        
        self.main_chart = ft.Container(
            content=ft.Text("Cargando datos...", text_align=ft.TextAlign.CENTER),
            height=300,
            border=ft.border.all(1, self.get_theme_colors()['border']),
            border_radius=10,
            padding=20,
            bgcolor=self.get_theme_colors()['bg_card']
        )
        
        # Tablas
        self.critical_items_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Item No")),
                ft.DataColumn(ft.Text("Descripción")),
                ft.DataColumn(ft.Text("Proyecto")),
                ft.DataColumn(ft.Text("Comprador")),
                ft.DataColumn(ft.Text("OH Total")),
                ft.DataColumn(ft.Text("Cobertura Actual")),
                ft.DataColumn(ft.Text("Cobertura c/PO")),
                ft.DataColumn(ft.Text("Estado")),
            ],
            rows=[]
        )
        
        self.po_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("PO No")),
                ft.DataColumn(ft.Text("Item No")),
                ft.DataColumn(ft.Text("Cantidad")),
                ft.DataColumn(ft.Text("Fecha Llegada")),
                ft.DataColumn(ft.Text("Días Restantes")),
                ft.DataColumn(ft.Text("Estado")),
            ],
            rows=[]
        )
        
        # Agregar controles a la página
        self.page.add(
            self.create_header(),
            self.create_filters_section(),
            ft.Divider(),
            self.create_kpis_section(),
            ft.Divider(),
            self.create_chart_section(),
            ft.Divider(),
            self.create_po_section(),
            ft.Divider(),
            self.create_critical_items_section()
        )
    
    def rebuild_ui(self):
        """Reconstruir la interfaz con el nuevo tema"""
        # Limpiar la página
        self.page.controls.clear()
        
        # Reconstruir todos los componentes
        self.build_initial_ui()
        
        # Actualizar con los datos actuales
        if self.analyzer:
            self.update_kpis()
            self.update_chart()
            self.update_po_table()
            self.update_critical_items_table()
    
    def create_header(self):
        """Crear header de la aplicación"""
        colors = self.get_theme_colors()
        
        return ft.Container(
            content=ft.Row([
                ft.Icon("analytics", size=40, color=colors['accent']),
                ft.Column([
                    ft.Text("Line of Balance Dashboard", 
                           size=24, 
                           weight=ft.FontWeight.BOLD, 
                           color=colors['text_primary']),
                    ft.Text("Sistema de Gestión de Inventarios - R4Database", 
                           size=14, 
                           color=colors['text_secondary'])
                ], spacing=0),
                ft.Container(expand=True),
                ft.Column([
                    ft.Row([
                        ft.Icon("circle", size=12),
                        self.connection_status_text
                    ], spacing=5),
                    ft.Row([
                        ft.Text("Última actualización:", color=colors['text_secondary']),
                        self.last_update_text
                    ], spacing=5)
                ], spacing=5),
                ft.Row([
                    self.theme_toggle_button,
                    ft.ElevatedButton(
                        "Actualizar",
                        icon="refresh",
                        on_click=self.manual_refresh
                    )
                ], spacing=10)
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=20,
            bgcolor=colors['bg_secondary'],
            border_radius=10
        )
    
    def create_filters_section(self):
        """Crear sección de filtros"""
        colors = self.get_theme_colors()
        
        return ft.Container(
            content=ft.Column([
                ft.Text("Filtros", 
                        size=18, 
                        weight=ft.FontWeight.BOLD, 
                        color=colors['text_primary']),
                ft.Row([
                    self.buyer_dropdown,
                    self.project_dropdown,
                    self.date_range_dropdown,
                ], spacing=20)
            ], spacing=10),
            padding=20,
            bgcolor=colors['bg_secondary'],
            border_radius=10
        )
    
    def create_kpi_card(self, title, value, icon, color):
        """Crear una tarjeta KPI individual"""
        colors = self.get_theme_colors()
        
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(icon, size=30, color=color),
                    ft.Container(expand=True),
                ]),
                ft.Text(value, size=24, weight=ft.FontWeight.BOLD, color=color),
                ft.Text(title, size=12, color=colors['text_secondary'])
            ], spacing=5),
            width=180,
            height=100,
            padding=15,
            bgcolor=colors['bg_card'],
            border_radius=10,
            border=ft.border.all(2, color)
        )
    
    def create_kpis_section(self):
        """Crear sección de KPIs ejecutiva"""
        colors = self.get_theme_colors()
        
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Indicadores Clave de Performance", 
                            size=20, 
                            weight=ft.FontWeight.W_600, 
                            color=colors['text_primary']),
                    ft.Container(expand=True),
                    ft.Container(
                        content=ft.Row([
                            ft.Icon("insights", size=16, color=colors['accent']),
                            ft.Text("Actualizado en tiempo real", 
                                    size=12, 
                                    color=colors['text_secondary'])
                        ], spacing=5),
                        padding=8,
                        bgcolor=colors['accent'] + "10",
                        border_radius=6
                    )
                ]),
                ft.Container(height=8),  # Espaciado
                self.kpi_cards
            ], spacing=0),
            padding=24,
            bgcolor=colors['bg_secondary'],
            border_radius=16,
            shadow=ft.BoxShadow(
                spread_radius=0,
                blur_radius=12,
                color="#00000008",
                offset=ft.Offset(0, 4)
            )
        )
    
    def create_chart_section(self):
        """Crear sección del gráfico principal"""
        colors = self.get_theme_colors()
        
        return ft.Container(
            content=ft.Column([
                ft.Text("Line of Balance - Proyección Simplificada", 
                        size=18, 
                        weight=ft.FontWeight.BOLD, 
                        color=colors['text_primary']),
                self.main_chart
            ], spacing=15),
            padding=20,
            bgcolor=colors['bg_secondary'],
            border_radius=10
        )
    
    def create_kpi_cards(self):
        """Crear tarjetas de KPIs"""
        return ft.Row([
            self.create_kpi_card("Items Totales", "0", "inventory_2", "#0000FF"),
            self.create_kpi_card("Items Críticos", "0", "warning", "#FF0000"),
            self.create_kpi_card("Cobertura Prom.", "0d", "timeline", "#00FF00"),
            self.create_kpi_card("Valor Inventario", "$0", "attach_money", "#800080"),
            self.create_kpi_card("POs Activas", "0", "shopping_cart", "#FFA500"),
        ], spacing=10, scroll=ft.ScrollMode.AUTO)
    
    def create_po_section(self):
        """Crear sección de Purchase Orders"""
        colors = self.get_theme_colors()
        
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Purchase Orders Activas", 
                            size=18, 
                            weight=ft.FontWeight.BOLD, 
                            color=colors['text_primary']),
                    ft.Container(expand=True),
                    ft.Row([
                        ft.Icon("local_shipping", color=colors['accent']),
                        ft.Text("Próximas Llegadas", 
                                size=14, 
                                color=colors['accent'])
                    ])
                ]),
                ft.Container(
                    content=self.po_table,
                    height=250,
                    border=ft.border.all(1, colors['border']),
                    border_radius=5,
                    padding=10,
                    bgcolor=colors['bg_card']
                )
            ], spacing=15),
            padding=20,
            bgcolor=colors['bg_secondary'],
            border_radius=10
        )
    
    def create_critical_items_section(self):
        """Crear sección de items críticos"""
        colors = self.get_theme_colors()
        
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Items que Requieren Atención", 
                            size=18, 
                            weight=ft.FontWeight.BOLD, 
                            color=colors['text_primary']),
                    ft.Container(expand=True),
                    ft.Icon("priority_high", color="#FF0000")
                ]),
                ft.Container(
                    content=self.critical_items_table,
                    height=300,
                    border=ft.border.all(1, colors['border']),
                    border_radius=5,
                    padding=10,
                    bgcolor=colors['bg_card']
                )
            ], spacing=15),
            padding=20,
            bgcolor=colors['bg_secondary'],
            border_radius=10
        )
    
    def refresh_data(self):
        """Actualizar datos desde la base de datos"""
        try:
            # Mostrar información de debug
            print(f"Intentando conectar a: {self.config.db_path}")
            
            # Verificar si el archivo existe
            import os
            if not os.path.exists(self.config.db_path):
                error_msg = f"❌ El archivo de base de datos no existe en: {self.config.db_path}"
                print(error_msg)
                self.show_error_dialog(error_msg)
                return
            
            print("✅ Archivo de base de datos encontrado")
            print("🔄 Conectando y obteniendo datos...")
            
            self.data, self.po_data = self.db_manager.connect_and_fetch_data()
            
            print(f"✅ Datos obtenidos: {len(self.data)} registros principales, {len(self.po_data)} POs")
            
            self.analyzer = LOBAnalyzer(self.data)
            
            self.update_connection_status()
            self.update_filters()
            self.update_kpis()
            self.update_chart()
            self.update_po_table()
            self.update_critical_items_table()
            
            print("✅ Dashboard actualizado correctamente")
            self.page.update()
            
        except Exception as e:
            error_msg = f"❌ Error: {str(e)}"
            print(error_msg)
            self.show_error_dialog(error_msg)
    
    def update_connection_status(self):
        """Actualizar estado de conexión"""
        if self.db_manager.connection_status == "connected":
            self.connection_status_text.value = "Conectado"
            self.connection_status_text.color = "#00FF00"
            if self.db_manager.last_update:
                self.last_update_text.value = self.db_manager.last_update.strftime("%H:%M:%S")
        elif self.db_manager.connection_status == "error":
            self.connection_status_text.value = "Error"
            self.connection_status_text.color = "#FF0000"
        else:
            self.connection_status_text.value = "Desconectado"
            self.connection_status_text.color = "#808080"
    
    def update_filters(self):
        """Actualizar opciones de filtros"""
        if not self.data.empty:
            buyers = ["Todos"] + sorted(self.data['AC'].dropna().unique().tolist())
            self.buyer_dropdown.options = [ft.dropdown.Option(b) for b in buyers]
            
            projects = ["Todos"] + sorted(self.data['Project'].dropna().unique().tolist())
            self.project_dropdown.options = [ft.dropdown.Option(p) for p in projects]
    
    def update_kpis(self):
        """Actualizar KPIs"""
        if self.analyzer:
            kpis = self.analyzer.get_kpis(self.selected_buyer, self.selected_project)
            
            cards = self.kpi_cards.controls
            if len(cards) >= 5:
                cards[0].content.controls[1].value = str(kpis['total_items'])
                cards[1].content.controls[1].value = str(kpis['critical_items'])
                cards[2].content.controls[1].value = f"{kpis['avg_coverage']}d"
                cards[3].content.controls[1].value = f"${kpis['total_inventory_value']/1000:.0f}K"
                cards[4].content.controls[1].value = str(kpis['active_pos'])
    
    def update_chart(self):
        """Actualizar gráfico principal"""
        if self.analyzer:
            filtered_data = self.analyzer.filter_data(self.selected_buyer, self.selected_project)
            chart = self.chart_manager.create_simple_chart(filtered_data)
            
            # Actualizar el contenedor del gráfico con colores del tema
            colors = self.get_theme_colors()
            self.main_chart.content = chart
            self.main_chart.bgcolor = colors['bg_card']
            self.main_chart.border = ft.border.all(1, colors['border'])
    
    def update_po_table(self):
        """Actualizar tabla de Purchase Orders"""
        if not hasattr(self.db_manager, 'po_timeline_data') or self.db_manager.po_timeline_data.empty:
            return
        
        po_data = self.db_manager.po_timeline_data.head(20)
        rows = []
        
        for _, po in po_data.iterrows():
            if po['PO_Status'] == 'overdue':
                status_color = "#FF0000"
                status_icon = "error"
            elif po['PO_Status'] == 'critical':
                status_color = "#FFA500"
                status_icon = "warning"
            else:
                status_color = "#00FF00"
                status_icon = "check_circle"
            
                rows.append(ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(po['PONo'])),
                        ft.DataCell(ft.Text(po['ItemNo'])),
                        ft.DataCell(ft.Text(f"{po['OpnQ']:,.0f}")),
                        ft.DataCell(ft.Text(po['RevPrDt'].strftime('%Y-%m-%d') if pd.notna(po['RevPrDt']) else 'N/A')),
                        ft.DataCell(ft.Text(f"{po['Days_to_Arrival']:+d}d" if pd.notna(po['Days_to_Arrival']) else 'N/A')),
                        ft.DataCell(ft.Icon(status_icon, color=status_color, size=16)),
                    ]
                ))
        
        self.po_table.rows = rows
    
    def update_critical_items_table(self):
        """Actualizar tabla de items críticos"""
        if self.analyzer:
            critical_items = self.analyzer.get_critical_items(self.selected_buyer, self.selected_project)
            rows = []
            
            for _, item in critical_items.iterrows():
                if item['Status'] == 'critical':
                    status_color = "#FF0000"
                else:
                    status_color = "#FFA500"
                
                next_po_date = "N/A"
                if pd.notna(item['Next_PO_Date']):
                    next_po_date = item['Next_PO_Date'].strftime('%m/%d')
                
                rows.append(ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(item['ItemNo'])),
                        ft.DataCell(ft.Text(item['Description'][:25] + "..." if len(item['Description']) > 25 else item['Description'])),
                        ft.DataCell(ft.Text(item['Project'])),
                        ft.DataCell(ft.Text(item['AC'])),
                        ft.DataCell(ft.Text(f"{item['OH_Total']:,.0f}")),
                        ft.DataCell(ft.Text(f"{item['Current_Coverage']:.1f}d", color=status_color)),
                        ft.DataCell(ft.Text(f"{item['Coverage']:.1f}d", color="#00FF00" if item['Coverage'] > item['Current_Coverage'] else status_color)),
                        ft.DataCell(ft.Icon("circle", color=status_color, size=12)),
                    ]
                ))
            
            self.critical_items_table.rows = rows[:15]
    
    def on_date_range_change(self, e):
        """Manejar cambio de rango de fechas"""
        self.config.date_range_days = int(self.date_range_dropdown.value)
        print(f"📅 Rango de fechas cambiado a: {self.config.date_range_days} días")
        
        # Recargar datos con el nuevo rango
        self.refresh_data()
        
    def on_filter_change(self, e):
        """Manejar cambios en filtros"""
        self.selected_buyer = self.buyer_dropdown.value
        self.selected_project = self.project_dropdown.value
        
        # Actualizar visualizaciones
        self.update_kpis()
        self.update_chart()
        self.update_po_table()
        self.update_critical_items_table()
        self.page.update()
    
    def manual_refresh(self, e):
        """Actualización manual"""
        self.refresh_data()
    
    def show_error_dialog(self, message):
        """Mostrar diálogo de error"""
        def close_error_dialog(e):
            error_dialog.open = False
            self.page.update()
            
        error_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Error"),
            content=ft.Text(message),
            actions=[
                ft.TextButton("OK", on_click=close_error_dialog)
            ],
        )
        
        self.page.dialog = error_dialog
        error_dialog.open = True
        self.page.update()

def main():
    """Punto de entrada principal"""
    app = LineOfBalanceApp()
    ft.app(target=app.main, view=ft.WEB_BROWSER, port=8550)

if __name__ == "__main__":
    main()