import flet as ft
import sqlite3
import pandas as pd
from datetime import datetime
import logging
import os

# Configuración de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Colores
class Colors:
    BLUE = "#2196F3"
    RED = "#F44336"
    GREEN = "#4CAF50"
    ORANGE = "#FF9800"
    PURPLE = "#9C27B0"
    GREY = "#9E9E9E"
    WHITE = "#FFFFFF"
    SURFACE_VARIANT = "#F5F5F5"
    LIGHT_GREEN = "#e8f5e8"
    LIGHT_RED = "#ffebee"
    LIGHT_ORANGE = "#fff3e0"
    DARK_BLUE = "#1976D2"
    TEAL = "#00796B"

class UltraFastDashboard:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Dashboard Ultra Rápido - Cobertura & Line of Balance"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.window_width = 1600
        self.page.window_height = 1000
        
        # Paths
        self.coverage_db_path = r"C:\Users\J.Vazquez\Desktop\MASTER_PLAN\Line_Of_Balance\CoberturaMaterialesV2.db"
        self.vendor_db_path = r"C:\Users\J.Vazquez\Desktop\MASTER_PLAN\Line_Of_Balance\CoberturaMaterialesV2.db"
        
        # Estado
        self.current_filters = {
            'process': "Todos",
            'vendor': "Todos", 
            'shortname': "Todos",
            'tactical': "Todos",
            'week': "Todas"
        }
        
        # LÍMITES PARA PERFORMANCE
        self.MAX_RECORDS = 5000  # Máximo 5K registros para UI
        self.DISPLAY_LIMIT = 100  # Mostrar solo 100 en listas
        self.LOB_LIMIT = 1000    # Line of Balance inicial
        
        self.summary_stats = None
        self.unique_stats = None
        self.sample_data = None
        self.line_of_balance_data = None
        self.weekly_balance_data = None  # NUEVO: Datos semanales con balance acumulativo
        self.lob_loaded_count = self.LOB_LIMIT
        
        self.setup_controls()
        self.setup_ui()
        self.load_available_filters()
    
    def setup_controls(self):
        """Controles básicos"""
        self.process_dropdown = ft.Dropdown(
            label="Proceso", width=180, on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todos")]
        )
        self.vendor_dropdown = ft.Dropdown(
            label="Vendor", width=180, on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todos")]
        )
        self.shortname_dropdown = ft.Dropdown(
            label="Shortname", width=180, on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todos")]
        )
        self.tactical_dropdown = ft.Dropdown(
            label="Buyer", width=180, on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todos")]
        )
        self.week_dropdown = ft.Dropdown(
            label="Semana", width=180, on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todas")]
        )
        
        self.refresh_button = ft.ElevatedButton(
            "⚡ Actualizar Rápido", on_click=self.ultra_fast_refresh,
            bgcolor=Colors.BLUE, color=Colors.WHITE
        )
        self.export_button = ft.ElevatedButton(
            "📊 Exportar Completo", on_click=self.export_full_data,
            bgcolor=Colors.GREEN, color=Colors.WHITE
        )
        self.load_more_lob_button = ft.ElevatedButton(
            "📈 Cargar Más LOB", on_click=self.load_more_line_of_balance,
            bgcolor=Colors.TEAL, color=Colors.WHITE, visible=False
        )
        
        self.status_text = ft.Text("Listo para actualización rápida...", color=Colors.GREY)
        
        # Contenedores
        self.kpi_row = ft.Row([], alignment=ft.MainAxisAlignment.SPACE_AROUND, wrap=True)
        self.unique_kpi_row = ft.Row([], alignment=ft.MainAxisAlignment.SPACE_AROUND, wrap=True)
        self.main_content = ft.Container(
            content=ft.Text("Presiona 'Actualizar Rápido' para cargar datos optimizados", 
                          size=16, text_align=ft.TextAlign.CENTER),
            padding=20, alignment=ft.alignment.center
        )
    
    def setup_ui(self):
        """UI optimizada y limpia"""
        header = ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Dashboard Ultra Rápido - Cobertura & LOB", size=24, weight=ft.FontWeight.BOLD),
                    ft.Row([self.refresh_button, self.export_button, self.load_more_lob_button], spacing=8),
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Divider(height=5, color=Colors.GREY),
                ft.Row([
                    ft.Text("Filtros:", size=14, weight=ft.FontWeight.BOLD),
                    self.process_dropdown, self.vendor_dropdown,
                    self.shortname_dropdown, self.tactical_dropdown, self.week_dropdown,
                ], spacing=8),
            ]),
            padding=15, bgcolor=Colors.SURFACE_VARIANT, border_radius=10, margin=10
        )
        
        status_bar = ft.Container(
            content=ft.Row([
                ft.Text("Estado:", color=Colors.GREY, size=12),
                self.status_text
            ]),
            padding=ft.padding.symmetric(horizontal=20, vertical=5)
        )
        
        self.page.add(
            ft.Column([
                header, 
                status_bar,
                ft.Container(content=self.main_content, expand=True, padding=10)
            ], expand=True, spacing=5)
        )
    
    def load_available_filters(self):
        """Cargar filtros rápido"""
        try:
            if not os.path.exists(self.vendor_db_path):
                self.status_text.value = "❌ Base R4Database no encontrada"
                self.status_text.color = Colors.RED
                self.page.update()
                return
            
            conn = sqlite3.connect(self.vendor_db_path)
            query = """
            SELECT DISTINCT Process, Vendor, Shortname, Tactical
            FROM Vendor 
            WHERE Process IS NOT NULL AND Vendor IS NOT NULL
            LIMIT 1000
            """
            filters_data = pd.read_sql_query(query, conn)
            conn.close()
            
            processes = ["Todos"] + sorted(filters_data['Process'].dropna().unique().tolist())[:50]
            vendors = ["Todos"] + sorted(filters_data['Vendor'].dropna().unique().tolist())[:100]
            shortnames = ["Todos"] + sorted(filters_data['Shortname'].dropna().unique().tolist())[:50]
            tacticals = ["Todos"] + sorted(filters_data['Tactical'].dropna().unique().tolist())[:50]
            
            self.process_dropdown.options = [ft.dropdown.Option(p) for p in processes]
            self.vendor_dropdown.options = [ft.dropdown.Option(v) for v in vendors]
            self.shortname_dropdown.options = [ft.dropdown.Option(s) for s in shortnames]
            self.tactical_dropdown.options = [ft.dropdown.Option(t) for t in tacticals]
            
            for dropdown in [self.process_dropdown, self.vendor_dropdown, self.shortname_dropdown, self.tactical_dropdown]:
                dropdown.value = "Todos"
            self.week_dropdown.value = "Todas"
            
            self.status_text.value = "✅ Filtros cargados rápidamente"
            self.status_text.color = Colors.GREEN
            
        except Exception as e:
            logger.error(f"Error cargando filtros: {e}")
            self.status_text.value = f"❌ Error: {str(e)}"
            self.status_text.color = Colors.RED
        
        self.page.update()
    
    def on_filter_change(self, e):
        """Cambio de filtros"""
        self.current_filters = {
            'process': self.process_dropdown.value or "Todos",
            'vendor': self.vendor_dropdown.value or "Todos",
            'shortname': self.shortname_dropdown.value or "Todos",
            'tactical': self.tactical_dropdown.value or "Todos",
            'week': self.week_dropdown.value or "Todas"
        }
        self.lob_loaded_count = self.LOB_LIMIT  # Reset LOB count on filter change
    
    def ultra_fast_refresh(self, e):
        """Actualización ultra rápida usando agregaciones SQL"""
        self.refresh_button.disabled = True
        self.status_text.value = "⚡ Cargando estadísticas agregadas y únicas..."
        self.status_text.color = Colors.ORANGE
        self.page.update()
        
        try:
            start_time = datetime.now()
            
            # PASO 1: Cargar estadísticas agregadas y únicas
            stats = self.load_aggregated_stats()
            unique_stats = self.load_unique_stats()
            
            if stats is None:
                self.show_alert("Error cargando estadísticas")
                return
            
            # PASO 2: Cargar muestra pequeña para listas
            sample = self.load_sample_data()
            
            # PASO 3: Cargar Line of Balance inicial
            lob_data = self.load_line_of_balance(self.LOB_LIMIT)
            
            # PASO 4: Cargar Weekly Balance (NUEVO)
            weekly_data = self.load_weekly_balance()
            
            # PASO 5: Cargar opciones de semanas para filtro
            if weekly_data is not None and len(weekly_data) > 0:
                self.load_week_filter_options(weekly_data)
            
            self.summary_stats = stats
            self.unique_stats = unique_stats
            self.sample_data = sample
            self.line_of_balance_data = lob_data
            self.weekly_balance_data = weekly_data
            self.lob_loaded_count = self.LOB_LIMIT
            
            # PASO 4: Actualizar UI
            self.update_kpis_fast()
            self.update_unique_kpis()
            self.create_fast_view()
            
            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            
            self.status_text.value = f"✅ Actualizado en {duration:.2f}s - Stats: {stats['total_records']:,} registros"
            self.status_text.color = Colors.GREEN
            self.load_more_lob_button.visible = True
            
        except Exception as ex:
            logger.error(f"Error: {ex}")
            self.status_text.value = f"❌ Error: {str(ex)}"
            self.status_text.color = Colors.RED
            self.show_alert(f"Error: {str(ex)}")
        
        finally:
            self.refresh_button.disabled = False
            self.page.update()
    
    def load_aggregated_stats(self):
        """Cargar estadísticas agregadas"""
        try:
            conn_vendor = sqlite3.connect(self.vendor_db_path)
            vendor_data = pd.read_sql_query(
                "SELECT DISTINCT Vendor, Process, Shortname, Tactical FROM Vendor", 
                conn_vendor
            )
            conn_vendor.close()
            
            conn = sqlite3.connect(self.coverage_db_path)
            
            where_conditions = ["1=1"]
            if self.current_filters['process'] != "Todos":
                where_conditions.append(f"v.Process = '{self.current_filters['process']}'")
            if self.current_filters['vendor'] != "Todos":
                where_conditions.append(f"v.Vendor = '{self.current_filters['vendor']}'")
            if self.current_filters['shortname'] != "Todos":
                where_conditions.append(f"v.Shortname = '{self.current_filters['shortname']}'")
            if self.current_filters['tactical'] != "Todos":
                where_conditions.append(f"v.Tactical = '{self.current_filters['tactical']}'")
            
            where_clause = " AND ".join(where_conditions)
            
            stats_query = f"""
            WITH filtered_data AS (
                SELECT c.Fill, c.Balance, c.ReqQty
                FROM Cobertura c
                LEFT JOIN (
                    SELECT DISTINCT Vendor, Process, Shortname, Tactical 
                    FROM Vendor
                ) v ON c.Vendor = v.Vendor
                WHERE {where_clause}
            )
            SELECT 
                COUNT(*) as total_records,
                SUM(CASE WHEN Fill = '❌ Faltante' THEN 1 ELSE 0 END) as faltantes,
                SUM(CASE WHEN Fill = 'OH' THEN 1 ELSE 0 END) as oh_covered,
                SUM(CASE WHEN Fill = 'PO' THEN 1 ELSE 0 END) as po_covered,
                SUM(CASE WHEN Fill = 'WO' THEN 1 ELSE 0 END) as wo_covered,
                SUM(CASE WHEN Fill = 'Rework' THEN 1 ELSE 0 END) as rework_covered,
                SUM(CASE WHEN Fill = '❌ Faltante' THEN ABS(Balance) ELSE 0 END) as total_faltante_qty,
                SUM(ReqQty) as total_req_qty,
                AVG(CASE WHEN Fill = '❌ Faltante' THEN ABS(Balance) ELSE NULL END) as avg_faltante
            FROM filtered_data
            """
            
            vendor_data.to_sql('temp_vendor', conn, if_exists='replace', index=False)
            stats_df = pd.read_sql_query(stats_query, conn)
            stats = stats_df.iloc[0].to_dict()
            
            conn.execute("DROP TABLE IF EXISTS temp_vendor")
            conn.close()
            
            return stats
            
        except Exception as e:
            logger.error(f"Error cargando estadísticas: {e}")
            if 'conn' in locals():
                conn.close()
            return None
    
    def load_unique_stats(self):
        """Cargar estadísticas de números únicos"""
        try:
            conn_vendor = sqlite3.connect(self.vendor_db_path)
            vendor_data = pd.read_sql_query(
                "SELECT DISTINCT Vendor, Process, Shortname, Tactical FROM Vendor", 
                conn_vendor
            )
            conn_vendor.close()
            
            conn = sqlite3.connect(self.coverage_db_path)
            
            where_conditions = ["1=1"]
            if self.current_filters['process'] != "Todos":
                where_conditions.append(f"v.Process = '{self.current_filters['process']}'")
            if self.current_filters['vendor'] != "Todos":
                where_conditions.append(f"v.Vendor = '{self.current_filters['vendor']}'")
            if self.current_filters['shortname'] != "Todos":
                where_conditions.append(f"v.Shortname = '{self.current_filters['shortname']}'")
            if self.current_filters['tactical'] != "Todos":
                where_conditions.append(f"v.Tactical = '{self.current_filters['tactical']}'")
            
            where_clause = " AND ".join(where_conditions)
            
            unique_stats_query = f"""
            WITH filtered_data AS (
                SELECT c.ItemNo, c.Vendor, c.Fill, v.Process, v.Shortname, v.Tactical
                FROM Cobertura c
                LEFT JOIN (
                    SELECT DISTINCT Vendor, Process, Shortname, Tactical 
                    FROM Vendor
                ) v ON c.Vendor = v.Vendor
                WHERE {where_clause}
            )
            SELECT 
                COUNT(DISTINCT ItemNo) as unique_items,
                COUNT(DISTINCT Vendor) as unique_vendors,
                COUNT(DISTINCT Process) as unique_processes,
                COUNT(DISTINCT Shortname) as unique_shortnames,
                COUNT(DISTINCT Tactical) as unique_buyers,
                COUNT(DISTINCT CASE WHEN Fill = '❌ Faltante' THEN ItemNo ELSE NULL END) as unique_faltante_items,
                COUNT(DISTINCT CASE WHEN Fill = 'OH' THEN ItemNo ELSE NULL END) as unique_oh_items,
                COUNT(DISTINCT CASE WHEN Fill = 'PO' THEN ItemNo ELSE NULL END) as unique_po_items
            FROM filtered_data
            """
            
            vendor_data.to_sql('temp_vendor', conn, if_exists='replace', index=False)
            unique_stats_df = pd.read_sql_query(unique_stats_query, conn)
            unique_stats = unique_stats_df.iloc[0].to_dict()
            
            conn.execute("DROP TABLE IF EXISTS temp_vendor")
            conn.close()
            
            return unique_stats
            
        except Exception as e:
            logger.error(f"Error cargando estadísticas únicas: {e}")
            if 'conn' in locals():
                conn.close()
            return None
    
    def load_sample_data(self):
        """Cargar muestra pequeña para mostrar en listas"""
        try:
            conn_vendor = sqlite3.connect(self.vendor_db_path)
            vendor_data = pd.read_sql_query(
                "SELECT DISTINCT Vendor, Process, Shortname, Tactical FROM Vendor", 
                conn_vendor
            )
            conn_vendor.close()
            
            conn = sqlite3.connect(self.coverage_db_path)
            
            where_conditions = ["1=1"]
            if self.current_filters['process'] != "Todos":
                where_conditions.append(f"v.Process = '{self.current_filters['process']}'")
            if self.current_filters['vendor'] != "Todos":
                where_conditions.append(f"v.Vendor = '{self.current_filters['vendor']}'")
            if self.current_filters['shortname'] != "Todos":
                where_conditions.append(f"v.Shortname = '{self.current_filters['shortname']}'")
            if self.current_filters['tactical'] != "Todos":
                where_conditions.append(f"v.Tactical = '{self.current_filters['tactical']}'")
            
            where_clause = " AND ".join(where_conditions)
            
            vendor_data.to_sql('temp_vendor', conn, if_exists='replace', index=False)
            
            sample_query = f"""
            SELECT c.ItemNo, c.ReqDate, c.ReqQty, c.Fill, c.Balance, c.Vendor,
                   v.Process, v.Shortname, v.Tactical
            FROM Cobertura c
            LEFT JOIN temp_vendor v ON c.Vendor = v.Vendor
            WHERE {where_clause} AND c.Fill = '❌ Faltante'
            ORDER BY ABS(c.Balance) DESC
            LIMIT {self.DISPLAY_LIMIT}
            """
            
            sample_df = pd.read_sql_query(sample_query, conn)
            
            conn.execute("DROP TABLE IF EXISTS temp_vendor")
            conn.close()
            
            return sample_df
            
        except Exception as e:
            logger.error(f"Error cargando muestra: {e}")
    def load_weekly_balance(self):
        """Cargar balance semanal acumulativo - TODAS las semanas del período"""
        try:
            conn_vendor = sqlite3.connect(self.vendor_db_path)
            vendor_data = pd.read_sql_query(
                "SELECT DISTINCT Vendor, Process, Shortname, Tactical FROM Vendor", 
                conn_vendor
            )
            conn_vendor.close()
            
            conn = sqlite3.connect(self.coverage_db_path)
            
            where_conditions = ["1=1"]
            if self.current_filters['process'] != "Todos":
                where_conditions.append(f"v.Process = '{self.current_filters['process']}'")
            if self.current_filters['vendor'] != "Todos":
                where_conditions.append(f"v.Vendor = '{self.current_filters['vendor']}'")
            if self.current_filters['shortname'] != "Todos":
                where_conditions.append(f"v.Shortname = '{self.current_filters['shortname']}'")
            if self.current_filters['tactical'] != "Todos":
                where_conditions.append(f"v.Tactical = '{self.current_filters['tactical']}'")
            
            where_clause = " AND ".join(where_conditions)
            
            vendor_data.to_sql('temp_vendor', conn, if_exists='replace', index=False)
            
            # Query para obtener datos semanales agrupados
            weekly_query = f"""
            WITH weekly_data AS (
                SELECT 
                    date(c.ReqDate, 'weekday 0', '-6 days') as week_start,
                    strftime('%Y-W%W', c.ReqDate) as week_label,
                    SUM(c.ReqQty) as req_qty,
                    SUM(CASE WHEN c.Fill = '❌ Faltante' THEN c.ReqQty ELSE 0 END) as faltante_qty,
                    SUM(CASE WHEN c.Fill = 'OH' THEN c.ReqQty ELSE 0 END) as oh_qty,
                    SUM(CASE WHEN c.Fill = 'PO' THEN c.ReqQty ELSE 0 END) as po_qty,
                    SUM(CASE WHEN c.Fill = 'WO' THEN c.ReqQty ELSE 0 END) as wo_qty,
                    COUNT(*) as item_count,
                    COUNT(DISTINCT c.ItemNo) as unique_items
                FROM Cobertura c
                LEFT JOIN temp_vendor v ON c.Vendor = v.Vendor
                WHERE {where_clause}
                GROUP BY week_start, week_label
            )
            SELECT *,
                   ROUND((req_qty - faltante_qty) * 100.0 / NULLIF(req_qty, 0), 1) as cobertura_pct
            FROM weekly_data
            ORDER BY week_start
            """
            
            weekly_df = pd.read_sql_query(weekly_query, conn)
            
            # Generar TODAS las semanas del período (past due + futuro)
            if len(weekly_df) > 0:
                min_date = pd.to_datetime(weekly_df['week_start'].min())
                max_date = pd.to_datetime(weekly_df['week_start'].max())
                
                # Extender rango: 4 semanas antes del mínimo, hasta el máximo + 12 semanas
                start_range = min_date - pd.Timedelta(weeks=4)
                end_range = max_date + pd.Timedelta(weeks=12)
                
                # Crear todas las semanas en el rango
                all_weeks = pd.date_range(start=start_range, end=end_range, freq='W-SUN')
                all_weeks_df = pd.DataFrame({
                    'week_start': all_weeks.strftime('%Y-%m-%d'),
                    'week_label': all_weeks.strftime('%Y-W%U'),
                    'req_qty': 0,
                    'faltante_qty': 0,
                    'oh_qty': 0,
                    'po_qty': 0,
                    'wo_qty': 0,
                    'item_count': 0,
                    'unique_items': 0,
                    'cobertura_pct': 100.0  # Sin requerimientos = 100% cubierto
                })
                
                # Merge con datos reales (reemplazar donde existan datos)
                weekly_df['week_start'] = pd.to_datetime(weekly_df['week_start']).dt.strftime('%Y-%m-%d')
                complete_weekly = all_weeks_df.set_index('week_start').combine_first(
                    weekly_df.set_index('week_start')
                ).reset_index()
                
                # Calcular balance acumulativo (ESTO ES CLAVE)
                complete_weekly = complete_weekly.sort_values('week_start')
                complete_weekly['balance_acumulativo'] = complete_weekly['faltante_qty'].cumsum()
                
                # Determinar status por semana
                def get_week_status(row):
                    if row['req_qty'] == 0:
                        return "Sin Requerimientos"
                    elif row['cobertura_pct'] >= 80:
                        return "Cubierta"
                    elif row['cobertura_pct'] >= 50:
                        return "Parcial"
                    else:
                        return "Crítica"
                
                complete_weekly['status'] = complete_weekly.apply(get_week_status, axis=1)
                
                # Determinar si es past due (comparar con fecha actual)
                today = datetime.now().date()
                complete_weekly['is_past_due'] = pd.to_datetime(complete_weekly['week_start']).dt.date < today
                
            else:
                complete_weekly = pd.DataFrame()
            
            conn.execute("DROP TABLE IF EXISTS temp_vendor")
            conn.close()
            
            return complete_weekly
            
        except Exception as e:
            logger.error(f"Error cargando balance semanal: {e}")
            return pd.DataFrame()
    
    def load_line_of_balance(self, limit):
        """Cargar datos para Line of Balance"""
        try:
            conn_vendor = sqlite3.connect(self.vendor_db_path)
            vendor_data = pd.read_sql_query(
                "SELECT DISTINCT Vendor, Process, Shortname, Tactical FROM Vendor", 
                conn_vendor
            )
            conn_vendor.close()
            
            conn = sqlite3.connect(self.coverage_db_path)
            
            where_conditions = ["1=1"]
            if self.current_filters['process'] != "Todos":
                where_conditions.append(f"v.Process = '{self.current_filters['process']}'")
            if self.current_filters['vendor'] != "Todos":
                where_conditions.append(f"v.Vendor = '{self.current_filters['vendor']}'")
            if self.current_filters['shortname'] != "Todos":
                where_conditions.append(f"v.Shortname = '{self.current_filters['shortname']}'")
            if self.current_filters['tactical'] != "Todos":
                where_conditions.append(f"v.Tactical = '{self.current_filters['tactical']}'")
            
            where_clause = " AND ".join(where_conditions)
            
            vendor_data.to_sql('temp_vendor', conn, if_exists='replace', index=False)
            
            lob_query = f"""
            SELECT c.ItemNo, c.ReqDate, c.ReqQty, c.Fill, c.Balance, c.Vendor,
                   v.Process, v.Shortname, v.Tactical,
                   CASE 
                       WHEN c.Fill = '❌ Faltante' THEN 'Crítico'
                       WHEN c.Fill = 'OH' THEN 'Cubierto'
                       WHEN c.Fill = 'PO' THEN 'En Proceso'
                       ELSE 'Otro'
                   END as Status
            FROM Cobertura c
            LEFT JOIN temp_vendor v ON c.Vendor = v.Vendor
            WHERE {where_clause}
            ORDER BY c.ReqDate, ABS(c.Balance) DESC
            LIMIT {limit}
            """
            
            lob_df = pd.read_sql_query(lob_query, conn)
            
            conn.execute("DROP TABLE IF EXISTS temp_vendor")
            conn.close()
            
            return lob_df
            
        except Exception as e:
            logger.error(f"Error cargando Line of Balance: {e}")
            return pd.DataFrame()
    
    def load_more_line_of_balance(self, e):
        """Cargar más datos para Line of Balance"""
        self.load_more_lob_button.disabled = True
        self.status_text.value = "📈 Cargando más datos LOB..."
        self.status_text.color = Colors.ORANGE
        self.page.update()
        
        try:
            new_limit = self.lob_loaded_count + 1000
            additional_data = self.load_line_of_balance(new_limit)
            
            if additional_data is not None and len(additional_data) > len(self.line_of_balance_data):
                self.line_of_balance_data = additional_data
                self.lob_loaded_count = new_limit
                
                # Actualizar solo la vista LOB
                self.create_fast_view()
                
                self.status_text.value = f"✅ LOB actualizado - {len(additional_data):,} registros cargados"
                self.status_text.color = Colors.GREEN
            else:
                self.status_text.value = "ℹ️ No hay más datos para cargar"
                self.status_text.color = Colors.GREY
            
        except Exception as ex:
            self.status_text.value = f"❌ Error cargando más LOB: {str(ex)}"
            self.status_text.color = Colors.RED
        
        finally:
            self.load_more_lob_button.disabled = False
            self.page.update()
    
    def update_kpis_fast(self):
        """KPIs totales en layout limpio"""
        if self.summary_stats is None:
            return
        
        stats = self.summary_stats
        
        # KPIs principales en una fila compacta
        main_kpis = ft.Row([
            self.create_main_kpi_card("Total", f"{stats['total_records']:,}", "📦", Colors.BLUE),
            self.create_main_kpi_card("Faltantes", f"{stats['faltantes']:,}", "❌", Colors.RED),
            self.create_main_kpi_card("OH", f"{stats['oh_covered']:,}", "✅", Colors.GREEN),
            self.create_main_kpi_card("PO", f"{stats['po_covered']:,}", "📋", Colors.ORANGE),
            self.create_main_kpi_card("WO", f"{stats['wo_covered']:,}", "🔧", Colors.BLUE),
        ], alignment=ft.MainAxisAlignment.SPACE_EVENLY, wrap=False, spacing=10)
        
        self.kpi_row.controls.clear()
        self.kpi_row.controls.append(main_kpis)
    
    def update_unique_kpis(self):
        """KPIs únicos en layout limpio"""
        if self.unique_stats is None:
            return
        
        stats = self.unique_stats
        
        # KPIs únicos en una fila compacta
        unique_kpis = ft.Row([
            self.create_unique_kpi_card("Items", f"{stats['unique_items']:,}", "🔢", Colors.TEAL),
            self.create_unique_kpi_card("Vendors", f"{stats['unique_vendors']:,}", "🏢", Colors.PURPLE),
            self.create_unique_kpi_card("Procesos", f"{stats['unique_processes']:,}", "⚙️", Colors.DARK_BLUE),
            self.create_unique_kpi_card("Faltantes", f"{stats['unique_faltante_items']:,}", "🔴", Colors.RED),
            self.create_unique_kpi_card("Buyers", f"{stats['unique_buyers']:,}", "👥", Colors.ORANGE),
        ], alignment=ft.MainAxisAlignment.SPACE_EVENLY, wrap=False, spacing=10)
        
        self.unique_kpi_row.controls.clear()
        self.unique_kpi_row.controls.append(unique_kpis)
    
    def create_main_kpi_card(self, title: str, value: str, emoji: str, color: str):
        """KPI card principal - más grande"""
        return ft.Container(
            content=ft.Column([
                ft.Text(emoji, size=24, color=color),
                ft.Text(title, size=12, color=Colors.GREY, weight=ft.FontWeight.BOLD),
                ft.Text(value, size=18, weight=ft.FontWeight.BOLD, color=color)
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            width=140, height=100, padding=10,
            bgcolor=Colors.WHITE, border_radius=10,
            border=ft.border.all(2, color), 
            shadow=ft.BoxShadow(spread_radius=1, blur_radius=3, color=Colors.GREY, offset=ft.Offset(0, 2))
        )
    
    def create_unique_kpi_card(self, title: str, value: str, emoji: str, color: str):
        """KPI card único - más pequeño"""
        return ft.Container(
            content=ft.Column([
                ft.Text(emoji, size=20, color=color),
                ft.Text(f"{title} únicos", size=10, color=Colors.GREY),
                ft.Text(value, size=14, weight=ft.FontWeight.BOLD, color=color)
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=3),
            width=120, height=80, padding=8,
            bgcolor=Colors.SURFACE_VARIANT, border_radius=8,
            border=ft.border.all(1, color)
        )
    
    def create_fast_view(self):
        """Vista principal con KPIs integrados"""
        if self.summary_stats is None:
            return
        
        # KPIs Section en la parte superior
        kpis_section = ft.Container(
            content=ft.Column([
                # Título y KPIs principales
                ft.Container(
                    content=ft.Column([
                        ft.Text("📊 Resumen General", size=18, weight=ft.FontWeight.BOLD, color=Colors.BLUE),
                        self.kpi_row,
                    ], spacing=10),
                    padding=15, bgcolor=Colors.WHITE, border_radius=10,
                    border=ft.border.all(1, Colors.BLUE)
                ),
                
                # KPIs únicos
                ft.Container(
                    content=ft.Column([
                        ft.Text("🔢 Métricas Únicas", size=16, weight=ft.FontWeight.BOLD, color=Colors.TEAL),
                        self.unique_kpi_row,
                    ], spacing=8),
                    padding=15, bgcolor=Colors.SURFACE_VARIANT, border_radius=10,
                    border=ft.border.all(1, Colors.TEAL)
                ),
            ], spacing=15),
            margin=10
        )
        
        # Tabs con contenido
        content_tabs = ft.Tabs(
            selected_index=0,
            tabs=[
                ft.Tab(text="📈 Distribución", content=self.create_distribution_view()),
                ft.Tab(text="📅 Balance Semanal", content=self.create_weekly_balance_view()),
                ft.Tab(text="❌ Top Faltantes", content=self.create_faltantes_fast()),
                ft.Tab(text="📊 Line of Balance", content=self.create_line_of_balance_view()),
                ft.Tab(text="📋 Estadísticas", content=self.create_stats_view())
            ],
            height=500
        )
        
        self.main_content.content = ft.Column([
            kpis_section,
            ft.Container(content=content_tabs, padding=10)
        ], spacing=0)
    
    def create_distribution_view(self):
        """Vista de distribución mejorada"""
        stats = self.summary_stats
        unique_stats = self.unique_stats
        
        total = stats['total_records']
        if total == 0:
            return ft.Text("No hay datos")
        
        faltantes_pct = (stats['faltantes'] / total) * 100
        oh_pct = (stats['oh_covered'] / total) * 100
        po_pct = (stats['po_covered'] / total) * 100
        wo_pct = (stats['wo_covered'] / total) * 100
        
        return ft.Column([
            ft.Text("📊 Distribución por Tipo de Cobertura", size=18, weight=ft.FontWeight.BOLD),
            
            # Barras visuales mejoradas
            ft.Container(
                content=ft.Column([
                    # OH
                    ft.Row([
                        ft.Container(width=max(200, int(oh_pct * 4)), height=30, bgcolor=Colors.GREEN, border_radius=5),
                        ft.Column([
                            ft.Text(f"ON HAND: {stats['oh_covered']:,} registros ({oh_pct:.1f}%)", 
                                   size=14, weight=ft.FontWeight.BOLD),
                            ft.Text(f"Items únicos: {unique_stats['unique_oh_items']:,}", 
                                   size=12, color=Colors.TEAL)
                        ], spacing=2)
                    ], spacing=15),
                    
                    # PO
                    ft.Row([
                        ft.Container(width=max(150, int(po_pct * 4)), height=30, bgcolor=Colors.ORANGE, border_radius=5),
                        ft.Column([
                            ft.Text(f"PURCHASE ORDERS: {stats['po_covered']:,} registros ({po_pct:.1f}%)", 
                                   size=14, weight=ft.FontWeight.BOLD),
                            ft.Text(f"Items únicos: {unique_stats['unique_po_items']:,}", 
                                   size=12, color=Colors.TEAL)
                        ], spacing=2)
                    ], spacing=15),
                    
                    # WO
                    ft.Row([
                        ft.Container(width=max(100, int(wo_pct * 4)), height=30, bgcolor=Colors.BLUE, border_radius=5),
                        ft.Column([
                            ft.Text(f"WORK ORDERS: {stats['wo_covered']:,} registros ({wo_pct:.1f}%)", 
                                   size=14, weight=ft.FontWeight.BOLD),
                        ], spacing=2)
                    ], spacing=15),
                    
                    # FALTANTES - Destacado
                    ft.Container(
                        content=ft.Row([
                            ft.Container(width=max(250, int(faltantes_pct * 4)), height=35, bgcolor=Colors.RED, border_radius=5),
                            ft.Column([
                                ft.Text(f"🚨 FALTANTES: {stats['faltantes']:,} registros ({faltantes_pct:.1f}%)", 
                                       size=16, weight=ft.FontWeight.BOLD, color=Colors.RED),
                                ft.Text(f"Items únicos faltantes: {unique_stats['unique_faltante_items']:,}", 
                                       size=12, color=Colors.RED, weight=ft.FontWeight.BOLD),
                                ft.Text(f"Cantidad total: {stats['total_faltante_qty']:,.0f} unidades", 
                                       size=12, color=Colors.RED)
                            ], spacing=2)
                        ], spacing=15),
                        padding=10, bgcolor=Colors.LIGHT_RED, border_radius=10,
                        border=ft.border.all(2, Colors.RED)
                    ),
                ], spacing=15),
                padding=20, bgcolor=Colors.WHITE, border_radius=10
            )
        ], scroll=ft.ScrollMode.AUTO)
    
    def create_weekly_balance_view(self):
        """Vista de Balance Semanal - Line of Balance Real"""
        if self.weekly_balance_data is None or len(self.weekly_balance_data) == 0:
            return ft.Container(
                content=ft.Text("No hay datos para Balance Semanal", size=16),
                padding=20, alignment=ft.alignment.center
            )
        
        data = self.weekly_balance_data
        
        # Header con información de período
        total_weeks = len(data)
        past_due_weeks = len(data[data['is_past_due'] == True])
        future_weeks = total_weeks - past_due_weeks
        
        header_info = ft.Container(
            content=ft.Column([
                ft.Text("📅 Balance Semanal - Line of Balance", size=18, weight=ft.FontWeight.BOLD),
                ft.Row([
                    ft.Text(f"Período: {total_weeks} semanas", size=12, color=Colors.GREY),
                    ft.Text(f"Past Due: {past_due_weeks}", size=12, color=Colors.RED),
                    ft.Text(f"Futuro: {future_weeks}", size=12, color=Colors.BLUE),
                ], spacing=20),
                ft.Text("💡 Arrastra el mouse sobre los números para copiar", size=10, color=Colors.GREY, italic=True)
            ], spacing=5),
            padding=10, bgcolor=Colors.SURFACE_VARIANT, border_radius=5
        )
        
        # Crear timeline semanal
        weekly_timeline = []
        
        for _, week in data.iterrows():
            # Determinar color por cobertura
            if week['status'] == "Sin Requerimientos":
                bar_color = Colors.GREY
                text_color = Colors.GREY
                status_emoji = "⚪"
            elif week['status'] == "Cubierta":
                bar_color = Colors.GREEN
                text_color = Colors.GREEN
                status_emoji = "✅"
            elif week['status'] == "Parcial":
                bar_color = Colors.ORANGE
                text_color = Colors.ORANGE
                status_emoji = "⚠️"
            else:  # Crítica
                bar_color = Colors.RED
                text_color = Colors.RED
                status_emoji = "🚨"
            
            # Determinar ancho de barra (máximo 300px)
            bar_width = max(20, min(300, int(week['cobertura_pct'] * 3)))
            
            # Determinar si es past due
            week_style = "normal"
            if week['is_past_due']:
                week_style = "past_due"
                text_color = Colors.RED
            
            # Crear fila de semana (SELECCIONABLE)
            week_row = ft.Container(
                content=ft.Column([
                    # Encabezado de semana
                    ft.Row([
                        ft.Container(
                            content=ft.Text(f"{week['week_label']}", size=11, weight=ft.FontWeight.BOLD),
                            width=80
                        ),
                        ft.Container(
                            content=ft.Text(week['week_start'], size=10, color=Colors.GREY),
                            width=100
                        ),
                        ft.Container(
                            content=ft.Text(f"{status_emoji} {week['status']}", size=10, color=text_color),
                            width=120
                        ),
                        ft.Container(
                            content=ft.Text("PAST DUE" if week['is_past_due'] else "", 
                                           size=9, color=Colors.RED, weight=ft.FontWeight.BOLD),
                            width=80
                        )
                    ], spacing=5),
                    
                    # Barra de progreso y métricas
                    ft.Row([
                        # Barra visual
                        ft.Container(width=bar_width, height=20, bgcolor=bar_color, border_radius=3),
                        
                        # Métricas (SELECCIONABLES)
                        ft.SelectionArea(
                            content=ft.Row([
                                ft.Text(f"{week['cobertura_pct']:.1f}%", size=12, weight=ft.FontWeight.BOLD, color=text_color),
                                ft.Text(f"Req: {week['req_qty']:.0f}", size=10),
                                ft.Text(f"Falt: {week['faltante_qty']:.0f}", size=10, color=Colors.RED),
                                ft.Text(f"Items: {week['unique_items']:.0f}", size=10, color=Colors.TEAL),
                                ft.Text(f"Bal.Acum: {week['balance_acumulativo']:.0f}", size=10, 
                                       color=Colors.RED if week['balance_acumulativo'] > 0 else Colors.GREEN,
                                       weight=ft.FontWeight.BOLD),
                            ], spacing=15)
                        )
                    ], spacing=10)
                ], spacing=3),
                padding=8, 
                bgcolor=Colors.LIGHT_RED if week['is_past_due'] else Colors.WHITE,
                border_radius=5,
                border=ft.border.all(1, Colors.RED if week['is_past_due'] else bar_color),
                margin=ft.margin.only(bottom=3),
                on_click=lambda e, w=week: self.show_week_details_with_items(w)  # Click para detalles
            )
            
            weekly_timeline.append(week_row)
        
        return ft.Column([
            header_info,
            ft.Container(
                content=ft.Column(weekly_timeline, scroll=ft.ScrollMode.AUTO),
                height=450, padding=10
            )
        ])
    
    def show_week_details_with_items(self, week_data):
        """Mostrar detalles de semana CON items específicos"""
        try:
            # Obtener items de la semana
            week_label = week_data['week_label']
            items_df = self.get_week_items_details(f"📅 {week_label}")
            
            if len(items_df) == 0:
                self.show_alert(f"No hay items para la semana {week_label}")
                return
            
            # Crear dialog con detalles + tabla de items
            def close_dialog(e):
                dialog.open = False
                self.page.update()
            
            # Resumen de la semana
            week_summary = ft.Container(
                content=ft.Column([
                    ft.Text(f"📅 Detalles de {week_label}", size=18, weight=ft.FontWeight.BOLD),
                    ft.Row([
                        ft.Text(f"📊 Cobertura: {week_data['cobertura_pct']:.1f}%", size=14),
                        ft.Text(f"📦 Items únicos: {len(items_df):,}", size=14),
                        ft.Text(f"⚠️ Faltantes: {len(items_df[items_df['Fill'] == '❌ Faltante']):,}", size=14, color=Colors.RED),
                    ], spacing=20),
                    ft.Text("🚨 PAST DUE" if week_data['is_past_due'] else "Semana futura", 
                           size=12, color=Colors.RED if week_data['is_past_due'] else Colors.GREEN)
                ], spacing=8),
                padding=15, bgcolor=Colors.SURFACE_VARIANT, border_radius=10
            )
            
            # Tabla de items
            items_table = []
            
            # Header tabla
            header = ft.Container(
                content=ft.Row([
                    ft.Text("Item", size=11, weight=ft.FontWeight.BOLD, width=120),
                    ft.Text("Fecha", size=11, weight=ft.FontWeight.BOLD, width=80),
                    ft.Text("Qty", size=11, weight=ft.FontWeight.BOLD, width=60),
                    ft.Text("Status", size=11, weight=ft.FontWeight.BOLD, width=80),
                    ft.Text("Balance", size=11, weight=ft.FontWeight.BOLD, width=70),
                    ft.Text("Vendor", size=11, weight=ft.FontWeight.BOLD, width=100),
                ], spacing=5),
                padding=8, bgcolor=Colors.BLUE, border_radius=5
            )
            items_table.append(header)
            
            # Filas de items (SELECCIONABLES)
            for _, item in items_df.head(100).iterrows():  # Máximo 100 items
                status_color = Colors.RED if item['status'] == 'Crítico' else (
                    Colors.GREEN if item['status'] == 'Cubierto' else Colors.ORANGE
                )
                
                item_row = ft.Container(
                    content=ft.SelectionArea(
                        content=ft.Row([
                            ft.Text(str(item['ItemNo']), size=10, width=120),
                            ft.Text(str(item['ReqDate'])[:10], size=10, width=80),
                            ft.Text(f"{item['ReqQty']:.0f}", size=10, width=60),
                            ft.Text(str(item['status']), size=10, width=80, color=status_color, weight=ft.FontWeight.BOLD),
                            ft.Text(f"{item['Balance']:.0f}", size=10, width=70, 
                                   color=Colors.RED if item['Balance'] < 0 else Colors.GREEN),
                            ft.Text(str(item['Vendor'])[:15], size=10, width=100),
                        ], spacing=5)
                    ),
                    padding=6, bgcolor=Colors.WHITE if item['status'] != 'Crítico' else Colors.LIGHT_RED,
                    border=ft.border.all(1, status_color), border_radius=3,
                    margin=ft.margin.only(bottom=1)
                )
                items_table.append(item_row)
            
            # Dialog principal
            dialog = ft.AlertDialog(
                modal=True,
                title=week_summary,
                content=ft.Container(
                    content=ft.Column([
                        ft.Text(f"📋 Items de la semana ({len(items_df):,} total, mostrando primeros 100)", 
                               size=12, color=Colors.GREY),
                        ft.Container(
                            content=ft.Column(items_table, scroll=ft.ScrollMode.AUTO),
                            height=400, width=800
                        )
                    ]),
                    width=900, height=500
                ),
                actions=[
                    ft.TextButton("Cerrar", on_click=close_dialog),
                    ft.ElevatedButton(f"Filtrar por esta semana", 
                                    on_click=lambda e: self.filter_by_week(week_label),
                                    bgcolor=Colors.BLUE, color=Colors.WHITE)
                ]
            )
            
            self.page.dialog = dialog
            dialog.open = True
            self.page.update()
            
        except Exception as e:
            logger.error(f"Error mostrando detalles de semana: {e}")
            self.show_alert(f"Error cargando detalles: {str(e)}")
    
    def filter_by_week(self, week_label):
        """Aplicar filtro por semana específica"""
        # Buscar la opción correcta en el dropdown
        for option in self.week_dropdown.options:
            if week_label in option.key:
                self.week_dropdown.value = option.key
                break
        
        # Cerrar dialog
        if self.page.dialog:
            self.page.dialog.open = False
        
        # Aplicar filtro y refrescar
        self.on_filter_change(None)
        self.ultra_fast_refresh(None)
        
        self.page.update()
    
    def show_week_details(self, week_data):
        """Mostrar detalles básicos de semana (método original simplificado)"""
        details = f"""📅 Detalles de {week_data['week_label']}

📊 Resumen:
• Fecha: {week_data['week_start']}
• Status: {week_data['status']}
• Cobertura: {week_data['cobertura_pct']:.1f}%

📦 Cantidades:
• Requerido: {week_data['req_qty']:.0f} unidades
• Faltante: {week_data['faltante_qty']:.0f} unidades
• OH: {week_data['oh_qty']:.0f} unidades
• PO: {week_data['po_qty']:.0f} unidades

📋 Items:
• Total registros: {week_data['item_count']:.0f}
• Items únicos: {week_data['unique_items']:.0f}
• Balance acumulativo: {week_data['balance_acumulativo']:.0f}

{"🚨 ESTA SEMANA ESTÁ VENCIDA (PAST DUE)" if week_data['is_past_due'] else ""}"""
        
        self.show_alert(details)
    
    def create_line_of_balance_view(self):
        """Vista de Line of Balance"""
        if self.line_of_balance_data is None or len(self.line_of_balance_data) == 0:
            return ft.Container(
                content=ft.Text("No hay datos para Line of Balance", size=16),
                padding=20, alignment=ft.alignment.center
            )
        
        # Header info
        header_info = ft.Container(
            content=ft.Row([
                ft.Text(f"📈 Line of Balance - {len(self.line_of_balance_data):,} registros", 
                       size=18, weight=ft.FontWeight.BOLD),
                ft.Text(f"(Mostrando primeros {self.lob_loaded_count:,})", 
                       size=12, color=Colors.GREY),
                ft.ElevatedButton(
                    "📊 Cargar +1000 más", 
                    on_click=self.load_more_line_of_balance,
                    bgcolor=Colors.TEAL, color=Colors.WHITE, scale=0.8
                )
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=10, bgcolor=Colors.SURFACE_VARIANT, border_radius=5
        )
        
        # Tabla LOB
        lob_table = []
        
        # Header tabla
        table_header = ft.Container(
            content=ft.Row([
                ft.Text("Item", size=11, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("Fecha Req", size=11, weight=ft.FontWeight.BOLD, width=80),
                ft.Text("Qty", size=11, weight=ft.FontWeight.BOLD, width=60),
                ft.Text("Status", size=11, weight=ft.FontWeight.BOLD, width=80),
                ft.Text("Balance", size=11, weight=ft.FontWeight.BOLD, width=70),
                ft.Text("Vendor", size=11, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("Proceso", size=11, weight=ft.FontWeight.BOLD, width=90),
            ], spacing=5),
            padding=8, bgcolor=Colors.BLUE, border_radius=5
        )
        lob_table.append(table_header)
        
        # Filas de datos
        for _, row in self.line_of_balance_data.head(50).iterrows():  # Mostrar solo 50 en UI
            status_color = Colors.RED if row['Status'] == 'Crítico' else (
                Colors.GREEN if row['Status'] == 'Cubierto' else Colors.ORANGE
            )
            
            row_container = ft.Container(
                content=ft.Row([
                    ft.Text(str(row['ItemNo'])[:15], size=10, width=100),
                    ft.Text(str(row['ReqDate'])[:10], size=10, width=80),
                    ft.Text(f"{row['ReqQty']:.0f}", size=10, width=60),
                    ft.Text(str(row['Status']), size=10, width=80, color=status_color, weight=ft.FontWeight.BOLD),
                    ft.Text(f"{row['Balance']:.0f}", size=10, width=70, 
                           color=Colors.RED if row['Balance'] < 0 else Colors.GREEN),
                    ft.Text(str(row['Vendor'])[:15], size=10, width=100),
                    ft.Text(str(row.get('Process', ''))[:12], size=10, width=90),
                ], spacing=5),
                padding=6, bgcolor=Colors.WHITE if row['Status'] != 'Crítico' else Colors.LIGHT_RED,
                border=ft.border.all(1, status_color), border_radius=3,
                margin=ft.margin.only(bottom=1)
            )
            lob_table.append(row_container)
        
        return ft.Column([
            header_info,
            ft.Container(
                content=ft.Column(lob_table, scroll=ft.ScrollMode.AUTO),
                height=500, padding=10
            )
        ])
    
    def create_summary_fast(self):
        """Resumen súper rápido"""
        stats = self.summary_stats
        unique_stats = self.unique_stats
        
        total = stats['total_records']
        if total == 0:
            return ft.Text("No hay datos")
        
        faltantes_pct = (stats['faltantes'] / total) * 100
        oh_pct = (stats['oh_covered'] / total) * 100
        po_pct = (stats['po_covered'] / total) * 100
        
        return ft.Column([
            ft.Text("📊 Resumen de Cobertura", size=18, weight=ft.FontWeight.BOLD),
            
            # Estadísticas generales
            ft.Container(
                content=ft.Column([
                    ft.Text("📦 Estadísticas Generales", size=16, weight=ft.FontWeight.BOLD, color=Colors.BLUE),
                    ft.Row([
                        ft.Text("Total registros:", size=14, weight=ft.FontWeight.BOLD),
                        ft.Text(f"{total:,}", size=14, color=Colors.BLUE)
                    ]),
                    ft.Row([
                        ft.Text("Items únicos:", size=14, weight=ft.FontWeight.BOLD),
                        ft.Text(f"{unique_stats['unique_items']:,}", size=14, color=Colors.TEAL)
                    ]),
                    ft.Row([
                        ft.Text("Vendors únicos:", size=14, weight=ft.FontWeight.BOLD),
                        ft.Text(f"{unique_stats['unique_vendors']:,}", size=14, color=Colors.PURPLE)
                    ]),
                ], spacing=8),
                padding=15, bgcolor=Colors.SURFACE_VARIANT, border_radius=10
            ),
            
            ft.Divider(),
            
            # Distribución visual
            ft.Text("📈 Distribución por Tipo de Cobertura", size=16, weight=ft.FontWeight.BOLD, color=Colors.GREEN),
            
            ft.Container(
                content=ft.Column([
                    # OH
                    ft.Row([
                        ft.Container(width=int(oh_pct * 3), height=25, bgcolor=Colors.GREEN, border_radius=5),
                        ft.Text(f"OH: {stats['oh_covered']:,} registros ({oh_pct:.1f}%)", size=12),
                        ft.Text(f"Items únicos: {unique_stats['unique_oh_items']:,}", size=11, color=Colors.TEAL)
                    ], spacing=10),
                    
                    # PO
                    ft.Row([
                        ft.Container(width=int(po_pct * 3), height=25, bgcolor=Colors.ORANGE, border_radius=5),
                        ft.Text(f"PO: {stats['po_covered']:,} registros ({po_pct:.1f}%)", size=12),
                        ft.Text(f"Items únicos: {unique_stats['unique_po_items']:,}", size=11, color=Colors.TEAL)
                    ], spacing=10),
                    
                    # WO
                    ft.Row([
                        ft.Container(width=100, height=25, bgcolor=Colors.BLUE, border_radius=5),
                        ft.Text(f"WO: {stats['wo_covered']:,} registros", size=12),
                    ], spacing=10),
                    
                    # Rework
                    ft.Row([
                        ft.Container(width=80, height=25, bgcolor=Colors.PURPLE, border_radius=5),
                        ft.Text(f"Rework: {stats['rework_covered']:,} registros", size=12),
                    ], spacing=10),
                    
                    # FALTANTES
                    ft.Row([
                        ft.Container(width=int(faltantes_pct * 3), height=25, bgcolor=Colors.RED, border_radius=5),
                        ft.Text(f"FALTANTES: {stats['faltantes']:,} registros ({faltantes_pct:.1f}%)", 
                               size=12, color=Colors.RED, weight=ft.FontWeight.BOLD),
                        ft.Text(f"Items únicos: {unique_stats['unique_faltante_items']:,}", 
                               size=11, color=Colors.RED, weight=ft.FontWeight.BOLD)
                    ], spacing=10)
                ], spacing=12),
                padding=20, bgcolor=Colors.SURFACE_VARIANT, border_radius=10
            ),
            
            ft.Divider(),
            
            # Métricas críticas
            ft.Container(
                content=ft.Column([
                    ft.Text("⚠️ Métricas Críticas", size=16, weight=ft.FontWeight.BOLD, color=Colors.RED),
                    ft.Row([
                        ft.Text("Cantidad total faltante:", size=14, weight=ft.FontWeight.BOLD),
                        ft.Text(f"{stats['total_faltante_qty']:,.0f} unidades", size=14, color=Colors.RED)
                    ]),
                    ft.Row([
                        ft.Text("Promedio por item faltante:", size=14, weight=ft.FontWeight.BOLD),
                        ft.Text(f"{stats.get('avg_faltante', 0):,.1f} unidades", size=14, color=Colors.RED)
                    ]),
                    ft.Row([
                        ft.Text("Items únicos faltantes:", size=14, weight=ft.FontWeight.BOLD),
                        ft.Text(f"{unique_stats['unique_faltante_items']:,} items", size=14, color=Colors.RED)
                    ])
                ], spacing=8),
                padding=15, bgcolor=Colors.LIGHT_RED, border_radius=10
            )
        ], scroll=ft.ScrollMode.AUTO)
    
    def create_faltantes_fast(self):
        """Top faltantes usando muestra pre-cargada"""
        if self.sample_data is None or len(self.sample_data) == 0:
            return ft.Container(
                content=ft.Text("✅ No hay items faltantes en la muestra", size=16, color=Colors.GREEN),
                padding=20, alignment=ft.alignment.center
            )
        
        items_list = []
        
        # Header
        header = ft.Container(
            content=ft.Row([
                ft.Text("Item", size=12, weight=ft.FontWeight.BOLD, width=120),
                ft.Text("Fecha", size=12, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("Req", size=12, weight=ft.FontWeight.BOLD, width=70),
                ft.Text("Faltante", size=12, weight=ft.FontWeight.BOLD, width=80),
                ft.Text("Vendor", size=12, weight=ft.FontWeight.BOLD, width=120),
                ft.Text("Proceso", size=12, weight=ft.FontWeight.BOLD, width=100),
            ], spacing=5),
            padding=10, bgcolor=Colors.SURFACE_VARIANT, border_radius=5
        )
        items_list.append(header)
        
        # Items (máximo 50 para performance)
        for _, item in self.sample_data.head(50).iterrows():
            row = ft.Container(
                content=ft.Row([
                    ft.Text(str(item['ItemNo']), size=11, width=120),
                    ft.Text(str(item['ReqDate'])[:10], size=11, width=100),
                    ft.Text(f"{item['ReqQty']:.0f}", size=11, width=70),
                    ft.Text(f"{abs(item['Balance']):.0f}", size=11, 
                           color=Colors.RED, weight=ft.FontWeight.BOLD, width=80),
                    ft.Text(str(item['Vendor']), size=11, width=120),
                    ft.Text(str(item.get('Process', '')), size=11, width=100),
                ], spacing=5),
                padding=8, bgcolor=Colors.LIGHT_RED,
                border=ft.border.all(1, Colors.RED), border_radius=5,
                margin=ft.margin.only(bottom=2)
            )
            items_list.append(row)
        
        return ft.Column([
            ft.Text(f"❌ Top {len(self.sample_data)} Items Más Críticos", 
                   size=18, weight=ft.FontWeight.BOLD, color=Colors.RED),
            ft.Text(f"Mostrando muestra de los más críticos de {self.summary_stats['faltantes']:,} total", 
                   size=12, color=Colors.GREY),
            ft.Column(items_list, scroll=ft.ScrollMode.AUTO, height=400)
        ])
    
    def create_stats_view(self):
        """Vista de estadísticas detalladas"""
        if self.summary_stats is None or self.unique_stats is None:
            return ft.Text("No hay estadísticas")
        
        stats = self.summary_stats
        unique_stats = self.unique_stats
        
        return ft.Column([
            ft.Text("📈 Estadísticas Detalladas", size=18, weight=ft.FontWeight.BOLD),
            
            # Estadísticas totales
            ft.Container(
                content=ft.Column([
                    ft.Text("📊 Resumen General", size=16, weight=ft.FontWeight.BOLD, color=Colors.BLUE),
                    ft.Row([ft.Text("Total registros:"), ft.Text(f"{stats['total_records']:,}", weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("Total cantidad requerida:"), ft.Text(f"{stats['total_req_qty']:,.0f}", weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("Items únicos:"), ft.Text(f"{unique_stats['unique_items']:,}", color=Colors.TEAL, weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("Vendors únicos:"), ft.Text(f"{unique_stats['unique_vendors']:,}", color=Colors.PURPLE, weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("Procesos únicos:"), ft.Text(f"{unique_stats['unique_processes']:,}", color=Colors.DARK_BLUE, weight=ft.FontWeight.BOLD)]),
                ], spacing=8),
                padding=15, bgcolor=Colors.SURFACE_VARIANT, border_radius=10
            ),
            
            ft.Divider(),
            
            # Cobertura por tipo
            ft.Container(
                content=ft.Column([
                    ft.Text("🏭 Cobertura por Tipo", size=16, weight=ft.FontWeight.BOLD, color=Colors.GREEN),
                    ft.Row([ft.Text("On Hand (OH):"), ft.Text(f"{stats['oh_covered']:,} registros", color=Colors.GREEN, weight=ft.FontWeight.BOLD), ft.Text(f"({unique_stats['unique_oh_items']:,} items únicos)", color=Colors.TEAL)]),
                    ft.Row([ft.Text("Purchase Orders (PO):"), ft.Text(f"{stats['po_covered']:,} registros", color=Colors.ORANGE, weight=ft.FontWeight.BOLD), ft.Text(f"({unique_stats['unique_po_items']:,} items únicos)", color=Colors.TEAL)]),
                    ft.Row([ft.Text("Work Orders (WO):"), ft.Text(f"{stats['wo_covered']:,} registros", color=Colors.BLUE, weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("Rework:"), ft.Text(f"{stats['rework_covered']:,} registros", color=Colors.PURPLE, weight=ft.FontWeight.BOLD)]),
                ], spacing=8),
                padding=15, bgcolor=Colors.LIGHT_GREEN, border_radius=10
            ),
            
            ft.Divider(),
            
            # Items faltantes
            ft.Container(
                content=ft.Column([
                    ft.Text("❌ Items Faltantes", size=16, weight=ft.FontWeight.BOLD, color=Colors.RED),
                    ft.Row([ft.Text("Registros faltantes:"), ft.Text(f"{stats['faltantes']:,}", color=Colors.RED, weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("Items únicos faltantes:"), ft.Text(f"{unique_stats['unique_faltante_items']:,}", color=Colors.RED, weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("Cantidad total faltante:"), ft.Text(f"{stats['total_faltante_qty']:,.0f} unidades", color=Colors.RED, weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("Promedio por item:"), ft.Text(f"{stats.get('avg_faltante', 0):,.1f} unidades", color=Colors.RED, weight=ft.FontWeight.BOLD)]),
                ], spacing=8),
                padding=15, bgcolor=Colors.LIGHT_RED, border_radius=10
            ),
            
            ft.Divider(),
            
            # Porcentajes
            ft.Container(
                content=ft.Column([
                    ft.Text("📊 Porcentajes", size=16, weight=ft.FontWeight.BOLD, color=Colors.GREY),
                    ft.Row([ft.Text("% Registros cubiertos:"), ft.Text(f"{((stats['total_records'] - stats['faltantes']) / stats['total_records'] * 100):,.1f}%", color=Colors.GREEN, weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("% Registros faltantes:"), ft.Text(f"{(stats['faltantes'] / stats['total_records'] * 100):,.1f}%", color=Colors.RED, weight=ft.FontWeight.BOLD)]),
                    ft.Row([ft.Text("% Items únicos faltantes:"), ft.Text(f"{(unique_stats['unique_faltante_items'] / unique_stats['unique_items'] * 100):,.1f}%", color=Colors.RED, weight=ft.FontWeight.BOLD)]),
                ], spacing=8),
                padding=15, bgcolor=Colors.SURFACE_VARIANT, border_radius=10
            )
        ], scroll=ft.ScrollMode.AUTO)
    
    def export_full_data(self, e):
        """Exportar datos completos (puede tomar tiempo)"""
        self.export_button.disabled = True
        self.status_text.value = "📊 Exportando datos completos..."
        self.status_text.color = Colors.ORANGE
        self.page.update()
        
        try:
            # Cargar datos completos para export
            conn_vendor = sqlite3.connect(self.vendor_db_path)
            vendor_data = pd.read_sql_query(
                "SELECT DISTINCT Vendor, Process, Shortname, Tactical FROM Vendor", 
                conn_vendor
            )
            conn_vendor.close()
            
            conn = sqlite3.connect(self.coverage_db_path)
            
            # Crear tabla temporal
            vendor_data.to_sql('temp_vendor', conn, if_exists='replace', index=False)
            
            # Construir WHERE clause
            where_conditions = ["1=1"]
            
            if self.current_filters['process'] != "Todos":
                where_conditions.append(f"v.Process = '{self.current_filters['process']}'")
            if self.current_filters['vendor'] != "Todos":
                where_conditions.append(f"v.Vendor = '{self.current_filters['vendor']}'")
            if self.current_filters['shortname'] != "Todos":
                where_conditions.append(f"v.Shortname = '{self.current_filters['shortname']}'")
            if self.current_filters['tactical'] != "Todos":
                where_conditions.append(f"v.Tactical = '{self.current_filters['tactical']}'")
            
            where_clause = " AND ".join(where_conditions)
            
            # Query completo
            full_query = f"""
            SELECT c.*, v.Process, v.Shortname, v.Tactical
            FROM Cobertura c
            LEFT JOIN temp_vendor v ON c.Vendor = v.Vendor
            WHERE {where_clause}
            ORDER BY c.ItemNo, c.ReqDate
            """
            
            # Cargar en chunks para evitar memoria
            chunk_size = 10000
            chunks = []
            
            for chunk in pd.read_sql_query(full_query, conn, chunksize=chunk_size):
                chunks.append(chunk)
            
            full_data = pd.concat(chunks, ignore_index=True)
            
            # Limpiar
            conn.execute("DROP TABLE IF EXISTS temp_vendor")
            conn.close()
            
            # Export
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filters_str = "_".join([f"{k}={v}" for k, v in self.current_filters.items() if v != "Todos"])
            if not filters_str:
                filters_str = "AllData"
            
            filename = f"Cobertura_Completa_{filters_str}_{timestamp}.csv"
            full_data.to_csv(filename, index=False)
            
            self.status_text.value = f"✅ Exportado: {filename}"
            self.status_text.color = Colors.GREEN
            
            self.show_alert(f"✅ Exportación completa:\n\n"
                          f"📁 Archivo: {filename}\n"
                          f"📦 Registros: {len(full_data):,}\n"
                          f"❌ Faltantes: {len(full_data[full_data['Fill'] == '❌ Faltante']):,}")
            
        except Exception as ex:
            self.status_text.value = f"❌ Error exportando"
            self.status_text.color = Colors.RED
            self.show_alert(f"❌ Error: {str(ex)}")
        
        finally:
            self.export_button.disabled = False
            self.page.update()
    
    def show_alert(self, message: str):
        """Alerta optimizada"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("Información"),
            content=ft.Text(message),
            actions=[ft.TextButton("OK", on_click=close_dlg)]
        )
        
        self.page.dialog = dlg
        dlg.open = True
        self.page.update()

def main(page: ft.Page):
    UltraFastDashboard(page)

if __name__ == "__main__":
    ft.app(target=main, assets_dir="assets")