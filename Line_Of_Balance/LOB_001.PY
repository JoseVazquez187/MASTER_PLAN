import flet as ft
import sqlite3
import pandas as pd
from datetime import datetime
import logging
import os

# Configuración de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Colores compatibles
class Colors:
    BLUE = "#2196F3"
    RED = "#F44336"
    GREEN = "#4CAF50"
    ORANGE = "#FF9800"
    PURPLE = "#9C27B0"
    GREY = "#9E9E9E"
    WHITE = "#FFFFFF"
    SURFACE_VARIANT = "#F5F5F5"
    GREY_300 = "#E0E0E0"

class OptimizedLOBDashboard:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Line of Balance - Dashboard Optimizado"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.window_width = 1400
        self.page.window_height = 900
        
        # Paths de bases de datos
        self.lob_db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\LOB_Database.db"
        
        # Estado de la aplicación
        self.current_filters = {
            'process': "Todos",
            'vendor': "Todos", 
            'shortname': "Todos",
            'tactical': "Todos"
        }
        
        # Datos cargados
        self.filtered_items = None
        self.lob_matrix_data = None
        
        # Crear controles
        self.setup_controls()
        self.setup_ui()
        
        # Auto-cargar filtros disponibles
        self.load_available_filters()
    
    def setup_controls(self):
        """Crear controles de la interfaz"""
        # Dropdowns
        self.process_dropdown = ft.Dropdown(
            label="Proceso",
            width=180,
            on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todos")]
        )
        
        self.vendor_dropdown = ft.Dropdown(
            label="Vendor",
            width=180,
            on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todos")]
        )
        
        self.shortname_dropdown = ft.Dropdown(
            label="Shortname",
            width=180,
            on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todos")]
        )
        
        self.tactical_dropdown = ft.Dropdown(
            label="Buyer (Tactical)",
            width=180,
            on_change=self.on_filter_change,
            options=[ft.dropdown.Option("Todos")]
        )
        
        # Botones
        self.refresh_button = ft.ElevatedButton(
            "🔄 Actualizar Vista",
            on_click=self.refresh_dashboard,
            bgcolor=Colors.BLUE,
            color=Colors.WHITE
        )
        
        self.regenerate_button = ft.ElevatedButton(
            "⚡ Regenerar LOB",
            on_click=self.regenerate_lob_data,
            bgcolor=Colors.ORANGE,
            color=Colors.WHITE
        )
        
        self.export_button = ft.ElevatedButton(
            "📊 Exportar Datos",
            on_click=self.export_filtered_data,
            bgcolor=Colors.GREEN,
            color=Colors.WHITE
        )
        
        # Status
        self.status_text = ft.Text("Cargando filtros disponibles...", color=Colors.GREY)
        self.last_update_text = ft.Text("", color=Colors.GREY, size=10)
        
        # Contenedores para pestañas
        self.kpi_container = ft.Row([], alignment=ft.MainAxisAlignment.SPACE_AROUND)
        self.dashboard_content = ft.Container(
            content=ft.Text("Selecciona filtros y presiona Actualizar Vista", 
                          size=16, text_align=ft.TextAlign.CENTER),
            padding=20,
            alignment=ft.alignment.center
        )
        self.lob_content = ft.Container(
            content=ft.Text("Line of Balance aparecerá aquí", size=16, text_align=ft.TextAlign.CENTER),
            padding=20,
            alignment=ft.alignment.center
        )
        self.critical_content = ft.Container(
            content=ft.Text("Items críticos aparecerán aquí", size=16, text_align=ft.TextAlign.CENTER),
            padding=20,
            alignment=ft.alignment.center
        )
        self.actions_content = ft.Container(
            content=ft.Text("Acciones PO aparecerán aquí", size=16, text_align=ft.TextAlign.CENTER),
            padding=20,
            alignment=ft.alignment.center
        )
    
    def setup_ui(self):
        """Configurar interfaz de usuario"""
        # Header
        header = ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Line of Balance - Dashboard Optimizado", size=24, weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.refresh_button,
                        self.regenerate_button,
                        self.export_button,
                    ], spacing=8),
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Divider(height=10),
                ft.Row([
                    ft.Text("Filtros:", size=16, weight=ft.FontWeight.BOLD),
                    self.process_dropdown,
                    self.vendor_dropdown,
                    self.shortname_dropdown,
                    self.tactical_dropdown,
                ], alignment=ft.MainAxisAlignment.START, spacing=8),
            ]),
            padding=20,
            bgcolor=Colors.SURFACE_VARIANT,
            border_radius=10
        )
        
        # Status bar
        status_bar = ft.Container(
            content=ft.Row([
                ft.Column([
                    ft.Row([
                        ft.Text("Estado:", color=Colors.GREY, size=12),
                        self.status_text
                    ]),
                    self.last_update_text
                ])
            ]),
            padding=ft.padding.symmetric(horizontal=20, vertical=10)
        )
        
        # Tabs para organizar contenido
        self.tabs = ft.Tabs(
            selected_index=0,
            tabs=[
                ft.Tab(
                    text="📊 Dashboard",
                    content=ft.Column([
                        self.kpi_container,
                        ft.Divider(height=20),
                        self.dashboard_content
                    ], scroll=ft.ScrollMode.AUTO)
                ),
                ft.Tab(
                    text="📈 Line of Balance",
                    content=self.lob_content
                ),
                ft.Tab(
                    text="⚠️ Items Críticos",
                    content=self.critical_content
                ),
                ft.Tab(
                    text="📋 Acciones PO",
                    content=self.actions_content
                )
            ]
        )
        
        # Layout principal
        self.page.add(
            ft.Column([
                header,
                status_bar,
                ft.Container(
                    content=self.tabs,
                    expand=True,
                    padding=20
                )
            ], expand=True)
        )
    
    def load_available_filters(self):
        """Cargar filtros disponibles desde la base LOB"""
        try:
            if not os.path.exists(self.lob_db_path):
                self.status_text.value = "❌ Base LOB no encontrada. Ejecuta el generador primero."
                self.status_text.color = Colors.RED
                self.page.update()
                return
            
            conn = sqlite3.connect(self.lob_db_path)
            
            # Cargar filtros únicos
            filters_query = """
            SELECT DISTINCT 
                process, vendor, shortname, tactical
            FROM lob_items 
            WHERE process IS NOT NULL AND vendor IS NOT NULL
            ORDER BY process, vendor, shortname, tactical
            """
            
            filters_data = pd.read_sql_query(filters_query, conn)
            
            # Obtener última actualización
            update_query = "SELECT MAX(last_updated) as last_update FROM lob_items"
            last_update = pd.read_sql_query(update_query, conn)['last_update'].iloc[0]
            
            conn.close()
            
            # Configurar dropdowns
            processes = ["Todos"] + sorted([str(p) for p in filters_data['process'].dropna().unique() if str(p) != 'nan'])
            self.process_dropdown.options = [ft.dropdown.Option(p) for p in processes]
            self.process_dropdown.value = "Todos"
            
            vendors = ["Todos"] + sorted([str(v) for v in filters_data['vendor'].dropna().unique() if str(v) != 'nan'])
            self.vendor_dropdown.options = [ft.dropdown.Option(v) for v in vendors]
            self.vendor_dropdown.value = "Todos"
            
            shortnames = ["Todos"] + sorted([str(s) for s in filters_data['shortname'].dropna().unique() if str(s) != 'nan'])
            self.shortname_dropdown.options = [ft.dropdown.Option(s) for s in shortnames]
            self.shortname_dropdown.value = "Todos"
            
            tacticals = ["Todos"] + sorted([str(t) for t in filters_data['tactical'].dropna().unique() if str(t) != 'nan'])
            self.tactical_dropdown.options = [ft.dropdown.Option(t) for t in tacticals]
            self.tactical_dropdown.value = "Todos"
            
            self.status_text.value = f"✅ Filtros cargados - Base LOB disponible"
            self.status_text.color = Colors.GREEN
            
            if last_update:
                self.last_update_text.value = f"Última actualización LOB: {last_update}"
            
            logger.info("Filtros disponibles cargados exitosamente")
            
        except Exception as e:
            logger.error(f"Error cargando filtros: {e}")
            self.status_text.value = f"❌ Error cargando filtros: {str(e)}"
            self.status_text.color = Colors.RED
        
        self.page.update()
    
    def on_filter_change(self, e):
        """Manejar cambios en filtros"""
        self.current_filters = {
            'process': self.process_dropdown.value or "Todos",
            'vendor': self.vendor_dropdown.value or "Todos",
            'shortname': self.shortname_dropdown.value or "Todos",
            'tactical': self.tactical_dropdown.value or "Todos"
        }
        logger.info(f"Filtros actualizados: {self.current_filters}")
    
    def refresh_dashboard(self, e):
        """Actualizar dashboard con filtros seleccionados"""
        logger.info("Actualizando dashboard...")
        
        self.refresh_button.disabled = True
        self.status_text.value = "📊 Cargando datos filtrados..."
        self.status_text.color = Colors.ORANGE
        self.page.update()
        
        try:
            # Cargar datos filtrados de la base LOB
            filtered_data = self.load_filtered_data()
            
            if filtered_data is None or len(filtered_data) == 0:
                self.show_alert("No se encontraron datos con los filtros seleccionados")
                self.status_text.value = "❌ No se encontraron datos"
                self.status_text.color = Colors.RED
                return
            
            logger.info(f"Datos filtrados cargados: {len(filtered_data)} items")
            self.filtered_items = filtered_data
            
            # Actualizar KPIs
            logger.info("Actualizando KPIs...")
            self.update_kpis()
            
            # Cargar proyecciones LOB
            logger.info("Cargando proyecciones LOB...")
            self.load_lob_projections()
            
            # Actualizar todas las pestañas
            logger.info("Actualizando pestañas...")
            self.update_all_tabs()
            
            self.status_text.value = f"✅ Dashboard actualizado - {len(filtered_data)} items cargados"
            self.status_text.color = Colors.GREEN
            
            logger.info("Dashboard actualizado exitosamente")
            
        except Exception as ex:
            logger.error(f"Error actualizando dashboard: {ex}", exc_info=True)
            self.status_text.value = f"❌ Error: {str(ex)}"
            self.status_text.color = Colors.RED
            self.show_alert(f"Error actualizando dashboard: {str(ex)}")
        
        finally:
            self.refresh_button.disabled = False
            self.page.update()
    
    def load_filtered_data(self):
        """Cargar datos filtrados desde la base LOB"""
        conn = sqlite3.connect(self.lob_db_path)
        
        # Construir WHERE clause
        where_conditions = ["1=1"]  # Condición base
        
        if self.current_filters['process'] != "Todos":
            where_conditions.append(f"process = '{self.current_filters['process']}'")
        
        if self.current_filters['vendor'] != "Todos":
            where_conditions.append(f"vendor = '{self.current_filters['vendor']}'")
        
        if self.current_filters['shortname'] != "Todos":
            where_conditions.append(f"shortname = '{self.current_filters['shortname']}'")
        
        if self.current_filters['tactical'] != "Todos":
            where_conditions.append(f"tactical = '{self.current_filters['tactical']}'")
        
        where_clause = " AND ".join(where_conditions)
        
        # Query optimizado
        query = f"""
        SELECT * FROM lob_items 
        WHERE {where_clause}
        ORDER BY current_coverage ASC, item_no
        """
        
        data = pd.read_sql_query(query, conn)
        conn.close()
        
        logger.info(f"Datos filtrados cargados: {len(data)} items")
        return data
    
    def load_lob_projections(self):
        """Cargar proyecciones LOB para items filtrados"""
        if self.filtered_items is None or len(self.filtered_items) == 0:
            return
        
        # Obtener lista de items filtrados (solo primeros 100 para performance)
        item_nos = self.filtered_items['item_no'].head(100).tolist()
        item_nos_str = "', '".join(item_nos)
        
        conn = sqlite3.connect(self.lob_db_path)
        
        projections_query = f"""
        SELECT * FROM lob_weekly_projection 
        WHERE item_no IN ('{item_nos_str}')
        ORDER BY item_no, week_number
        """
        
        self.lob_matrix_data = pd.read_sql_query(projections_query, conn)
        conn.close()
        
        logger.info(f"Proyecciones LOB cargadas: {len(self.lob_matrix_data)} registros")
    
    def update_kpis(self):
        """Actualizar tarjetas KPI"""
        if self.filtered_items is None:
            return
        
        total_items = len(self.filtered_items)
        critical_items = len(self.filtered_items[self.filtered_items['status'] == 'critical'])
        warning_items = len(self.filtered_items[self.filtered_items['status'] == 'warning'])
        avg_coverage = self.filtered_items['current_coverage'].mean()
        total_value = self.filtered_items['inventory_value'].sum()
        active_pos = len(self.filtered_items[self.filtered_items['open_po_qty'] > 0])
        
        self.kpi_container.controls.clear()
        self.kpi_container.controls.extend([
            self.create_kpi_card("Total Items", str(total_items), "📦", Colors.BLUE),
            self.create_kpi_card("Items Críticos", str(critical_items), "⚠️", Colors.RED),
            self.create_kpi_card("Items Warning", str(warning_items), "🟡", Colors.ORANGE),
            self.create_kpi_card("Cobertura Promedio", f"{avg_coverage:.1f} días", "⏰", Colors.GREEN),
            self.create_kpi_card("Valor Inventario", f"${total_value:,.0f}", "💰", Colors.PURPLE),
            self.create_kpi_card("POs Activas", str(active_pos), "🛒", Colors.ORANGE)
        ])
    
    def update_all_tabs(self):
        """Actualizar contenido de todas las pestañas"""
        try:
            logger.info("Iniciando actualización de pestañas...")
            
            # Pestaña Dashboard - Resumen
            self.update_dashboard_summary()
            
            # Pestaña Line of Balance - Matriz
            self.update_lob_matrix_tab()
            
            # Pestaña Items Críticos
            self.update_critical_items_tab()
            
            # Pestaña Acciones PO
            self.update_po_actions_tab()
            
            logger.info("Todas las pestañas actualizadas exitosamente")
            self.page.update()
            
        except Exception as e:
            logger.error(f"Error actualizando pestañas: {e}", exc_info=True)
            self.show_alert(f"Error actualizando pestañas: {str(e)}")
    
    def update_dashboard_summary(self):
        """Actualizar resumen en dashboard"""
        try:
            if self.filtered_items is None:
                return
            
            # Análisis por estado
            status_counts = self.filtered_items['status'].value_counts()
            
            # Top 10 items más críticos
            top_critical = self.filtered_items.nsmallest(10, 'current_coverage')
            
            summary_content = ft.Column([
                ft.Text("📊 Resumen de Estado de Items", size=18, weight=ft.FontWeight.BOLD),
                
                # Gráfico de estado
                ft.Row([
                    ft.Container(
                        content=ft.Column([
                            ft.Text("🔴", size=30),
                            ft.Text("Críticos", size=12),
                            ft.Text(str(status_counts.get('critical', 0)), 
                                   size=20, weight=ft.FontWeight.BOLD, color=Colors.RED)
                        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                        width=120, height=80, padding=10,
                        bgcolor=Colors.SURFACE_VARIANT, border_radius=10
                    ),
                    ft.Container(
                        content=ft.Column([
                            ft.Text("🟡", size=30),
                            ft.Text("Advertencia", size=12),
                            ft.Text(str(status_counts.get('warning', 0)), 
                                   size=20, weight=ft.FontWeight.BOLD, color=Colors.ORANGE)
                        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                        width=120, height=80, padding=10,
                        bgcolor=Colors.SURFACE_VARIANT, border_radius=10
                    ),
                    ft.Container(
                        content=ft.Column([
                            ft.Text("🟢", size=30),
                            ft.Text("Seguros", size=12),
                            ft.Text(str(status_counts.get('safe', 0)), 
                                   size=20, weight=ft.FontWeight.BOLD, color=Colors.GREEN)
                        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                        width=120, height=80, padding=10,
                        bgcolor=Colors.SURFACE_VARIANT, border_radius=10
                    )
                ], alignment=ft.MainAxisAlignment.SPACE_AROUND),
                
                ft.Divider(height=20),
                
                # Top 10 más críticos como lista simple
                ft.Text("🚨 Top 10 Items Más Críticos", size=16, weight=ft.FontWeight.BOLD, color=Colors.RED),
                self.create_critical_items_list(top_critical)
            ])
            
            self.dashboard_content.content = summary_content
            
        except Exception as e:
            logger.error(f"Error en update_dashboard_summary: {e}", exc_info=True)
            self.dashboard_content.content = ft.Text(f"Error cargando resumen: {str(e)}", color=Colors.RED)
    
    def create_critical_items_list(self, critical_items):
        """Crear lista simple de items críticos"""
        if len(critical_items) == 0:
            return ft.Text("No hay items críticos", color=Colors.GREEN)
        
        items_list = []
        
        # Header
        header = ft.Container(
            content=ft.Row([
                ft.Text("Item", size=12, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("Descripción", size=12, weight=ft.FontWeight.BOLD, width=250),
                ft.Text("Cobertura", size=12, weight=ft.FontWeight.BOLD, width=80),
                ft.Text("OH", size=12, weight=ft.FontWeight.BOLD, width=70),
                ft.Text("Vendor", size=12, weight=ft.FontWeight.BOLD, width=100),
            ], spacing=5),
            padding=10,
            bgcolor=Colors.SURFACE_VARIANT,
            border_radius=5
        )
        items_list.append(header)
        
        # Items
        for _, item in critical_items.head(10).iterrows():
            item_row = ft.Container(
                content=ft.Row([
                    ft.Text(str(item['item_no']), size=11, width=100),
                    ft.Text(str(item['description'])[:30] + "..." if len(str(item['description'])) > 30 else str(item['description']), 
                           size=11, width=250),
                    ft.Text(f"{item['current_coverage']:.1f}d", size=11, color=Colors.RED, weight=ft.FontWeight.BOLD, width=80),
                    ft.Text(f"{item['total_oh']:.0f}", size=11, width=70),
                    ft.Text(str(item['vendor']), size=11, width=100),
                ], spacing=5),
                padding=8,
                border=ft.border.all(1, Colors.GREY_300),
                border_radius=5,
                margin=ft.margin.only(bottom=2)
            )
            items_list.append(item_row)
        
        return ft.Column(items_list, height=350, scroll=ft.ScrollMode.AUTO)
    
    def update_lob_matrix_tab(self):
        """Actualizar pestaña Line of Balance"""
        try:
            if self.lob_matrix_data is None or len(self.lob_matrix_data) == 0:
                self.lob_content.content = ft.Column([
                    ft.Text("📈 LINE OF BALANCE", size=18, weight=ft.FontWeight.BOLD, color=Colors.BLUE),
                    ft.Text("No hay datos de proyección LOB disponibles", size=14, color=Colors.GREY),
                ])
                return
            
            # Crear vista LOB simplificada
            lob_view = self.create_lob_view()
            
            matrix_content = ft.Column([
                ft.Text("📈 LINE OF BALANCE - Proyección", size=18, weight=ft.FontWeight.BOLD, color=Colors.BLUE),
                ft.Text(f"Proyección para {len(self.filtered_items)} items (mostrando primeros 20)", size=12, color=Colors.GREY),
                lob_view,
                self.create_legend()
            ], scroll=ft.ScrollMode.AUTO)
            
            self.lob_content.content = matrix_content
            
        except Exception as e:
            logger.error(f"Error en update_lob_matrix_tab: {e}", exc_info=True)
            self.lob_content.content = ft.Text(f"Error cargando LOB: {str(e)}", color=Colors.RED)
    
    def create_lob_view(self):
        """Crear vista LOB simplificada"""
        # Resumir por item
        items_summary = self.lob_matrix_data.groupby('item_no').agg({
            'projected_inventory': 'min',
            'coverage_days': 'min',
            'status_color': lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else 'green',
            'action_required': lambda x: ', '.join(x.unique())
        }).reset_index()
        
        lob_rows = []
        
        # Header
        header = ft.Container(
            content=ft.Row([
                ft.Text("Item", size=12, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("OH Actual", size=12, weight=ft.FontWeight.BOLD, width=80),
                ft.Text("Min Proyectado", size=12, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("Min Cobertura", size=12, weight=ft.FontWeight.BOLD, width=90),
                ft.Text("Estado", size=12, weight=ft.FontWeight.BOLD, width=80),
                ft.Text("Acciones", size=12, weight=ft.FontWeight.BOLD, width=150),
            ], spacing=5),
            padding=10,
            bgcolor=Colors.SURFACE_VARIANT,
            border_radius=5
        )
        lob_rows.append(header)
        
        # Datos
        for _, item in items_summary.head(20).iterrows():
            # Obtener datos del item
            item_data = self.filtered_items[self.filtered_items['item_no'] == item['item_no']]
            if len(item_data) == 0:
                continue
                
            item_info = item_data.iloc[0]
            
            # Color según estado
            status_color = item['status_color']
            if status_color == 'red':
                bg_color = "#ffebee"
            elif status_color == 'orange':  
                bg_color = "#fff3e0"
            else:
                bg_color = "#e8f5e8"
            
            row = ft.Container(
                content=ft.Row([
                    ft.Text(str(item['item_no']), size=11, width=100),
                    ft.Text(f"{item_info['total_oh']:.0f}", size=11, width=80),
                    ft.Text(f"{item['projected_inventory']:.0f}", size=11, width=100),
                    ft.Text(f"{item['coverage_days']:.1f}d", size=11, width=90),
                    ft.Text(item_info['status'].upper(), size=11, width=80),
                    ft.Text(item['action_required'][:20] + "..." if len(item['action_required']) > 20 else item['action_required'], 
                           size=10, width=150),
                ], spacing=5),
                padding=8,
                bgcolor=bg_color,
                border=ft.border.all(1, Colors.GREY_300),
                border_radius=5,
                margin=ft.margin.only(bottom=2)
            )
            lob_rows.append(row)
        
        return ft.Column(lob_rows, height=400, scroll=ft.ScrollMode.AUTO)
    
    def update_critical_items_tab(self):
        """Actualizar pestaña de items críticos"""
        try:
            if self.filtered_items is None:
                return
            
            critical_items = self.filtered_items[self.filtered_items['status'] == 'critical']
            
            if len(critical_items) == 0:
                content = ft.Container(
                    content=ft.Text("✅ No hay items críticos", size=16, color=Colors.GREEN),
                    padding=20,
                    alignment=ft.alignment.center
                )
            else:
                content = ft.Column([
                    ft.Text(f"⚠️ Items Críticos ({len(critical_items)} items)", 
                           size=18, weight=ft.FontWeight.BOLD, color=Colors.RED),
                    ft.Text("Items con cobertura menor a 5 días", size=12, color=Colors.GREY),
                    self.create_critical_items_list(critical_items)
                ], scroll=ft.ScrollMode.AUTO)
            
            self.critical_content.content = content
            
        except Exception as e:
            logger.error(f"Error en update_critical_items_tab: {e}")
            self.critical_content.content = ft.Text(f"Error: {str(e)}", color=Colors.RED)
    
    def update_po_actions_tab(self):
        """Actualizar pestaña de acciones PO"""
        try:
            if self.lob_matrix_data is None:
                content = ft.Container(
                    content=ft.Text("No hay datos de proyección", size=14, color=Colors.GREY),
                    padding=20
                )
                self.actions_content.content = content
                return
            
            # Obtener acciones requeridas
            actions_needed = self.lob_matrix_data[
                self.lob_matrix_data['action_required'].isin(['place_po', 'urgent_po'])
            ].copy()
            
            if len(actions_needed) == 0:
                content = ft.Container(
                    content=ft.Text("✅ No hay acciones PO urgentes", size=16, color=Colors.GREEN),
                    padding=20
                )
            else:
                content = ft.Column([
                    ft.Text(f"📋 Acciones PO ({len(actions_needed)} acciones)", 
                           size=18, weight=ft.FontWeight.BOLD, color=Colors.ORANGE),
                    
                    # Resumen
                    ft.Row([
                        ft.Container(
                            content=ft.Column([
                                ft.Text("🚨", size=24),
                                ft.Text("URGENT PO", size=10, weight=ft.FontWeight.BOLD),
                                ft.Text(str(len(actions_needed[actions_needed['action_required'] == 'urgent_po'])), 
                                       size=16, color=Colors.RED)
                            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                            width=100, height=70, padding=10,
                            bgcolor=Colors.SURFACE_VARIANT, border_radius=10
                        ),
                        ft.Container(
                            content=ft.Column([
                                ft.Text("📝", size=24),
                                ft.Text("PLACE PO", size=10, weight=ft.FontWeight.BOLD),
                                ft.Text(str(len(actions_needed[actions_needed['action_required'] == 'place_po'])), 
                                       size=16, color=Colors.ORANGE)
                            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                            width=100, height=70, padding=10,
                            bgcolor=Colors.SURFACE_VARIANT, border_radius=10
                        )
                    ], spacing=10),
                    
                    ft.Divider(),
                    self.create_actions_list(actions_needed)
                ], scroll=ft.ScrollMode.AUTO)
            
            self.actions_content.content = content
            
        except Exception as e:
            logger.error(f"Error en update_po_actions_tab: {e}")
            self.actions_content.content = ft.Text(f"Error: {str(e)}", color=Colors.RED)
    
    def create_actions_list(self, actions_data):
        """Crear lista de acciones PO"""
        action_rows = []
        
        # Header
        header = ft.Container(
            content=ft.Row([
                ft.Text("Item", size=12, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("Semana", size=12, weight=ft.FontWeight.BOLD, width=80),
                ft.Text("Acción", size=12, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("Inv. Proyectado", size=12, weight=ft.FontWeight.BOLD, width=100),
                ft.Text("Cobertura", size=12, weight=ft.FontWeight.BOLD, width=80),
                ft.Text("Urgencia", size=12, weight=ft.FontWeight.BOLD, width=80),
            ], spacing=5),
            padding=10,
            bgcolor=Colors.SURFACE_VARIANT,
            border_radius=5
        )
        action_rows.append(header)
        
        # Datos
        for _, action in actions_data.head(30).iterrows():
            urgency_color = Colors.RED if action['action_required'] == 'urgent_po' else Colors.ORANGE
            urgency_text = "🔴 ALTA" if action['action_required'] == 'urgent_po' else "🟡 MEDIA"
            
            row = ft.Container(
                content=ft.Row([
                    ft.Text(str(action['item_no']), size=11, width=100),
                    ft.Text(str(action['week_label']), size=11, width=80),
                    ft.Text(action['action_required'].upper().replace('_', ' '), 
                           size=11, color=urgency_color, weight=ft.FontWeight.BOLD, width=100),
                    ft.Text(f"{action['projected_inventory']:.0f}", size=11, width=100),
                    ft.Text(f"{action['coverage_days']:.1f}d", size=11, width=80),
                    ft.Text(urgency_text, size=11, width=80),
                ], spacing=5),
                padding=8,
                border=ft.border.all(1, Colors.GREY_300),
                border_radius=5,
                margin=ft.margin.only(bottom=2)
            )
            action_rows.append(row)
        
        return ft.Column(action_rows, height=350, scroll=ft.ScrollMode.AUTO)
    
    def create_legend(self):
        """Crear leyenda"""
        return ft.Container(
            content=ft.Column([
                ft.Text("📖 Leyenda:", weight=ft.FontWeight.BOLD),
                ft.Row([
                    ft.Text("🟢 SAFE (>15d)", color=Colors.GREEN, size=12),
                    ft.Text("🟡 WARNING (5-15d)", color=Colors.ORANGE, size=12),
                    ft.Text("🔴 CRITICAL (<5d)", color=Colors.RED, size=12),
                ], spacing=15),
                ft.Row([
                    ft.Text("📝 PLACE PO", color=Colors.ORANGE, size=12),
                    ft.Text("🚨 URGENT PO", color=Colors.RED, size=12),
                    ft.Text("✅ OK", color=Colors.GREY, size=12)
                ], spacing=15)
            ]),
            padding=15,
            bgcolor=Colors.SURFACE_VARIANT,
            border_radius=5,
            margin=ft.margin.only(top=10)
        )
    
    def create_kpi_card(self, title: str, value: str, emoji: str, color: str):
        """Crear tarjeta KPI"""
        return ft.Container(
            content=ft.Column([
                ft.Text(emoji, size=30, color=color),
                ft.Text(title, size=12, color=Colors.GREY),
                ft.Text(value, size=20, weight=ft.FontWeight.BOLD)
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            width=180,
            height=120,
            padding=15,
            bgcolor=Colors.SURFACE_VARIANT,
            border_radius=10
        )
    
    def regenerate_lob_data(self, e):
        """Regenerar datos LOB"""
        self.regenerate_button.disabled = True
        self.status_text.value = "⚡ Regenerando datos LOB..."
        self.status_text.color = Colors.ORANGE
        self.page.update()
        
        try:
            import subprocess
            import sys
            
            # Buscar generador
            possible_names = ["lob_generator.py", "LOB_V2.py"]
            generator_path = None
            
            for name in possible_names:
                if os.path.exists(name):
                    generator_path = name
                    break
            
            if not generator_path:
                self.show_alert("❌ No se encontró el generador LOB")
                return
            
            # Ejecutar
            result = subprocess.run([sys.executable, generator_path], 
                                  capture_output=True, text=True, timeout=1800)
            
            if result.returncode == 0:
                self.status_text.value = "✅ LOB regenerado exitosamente"
                self.status_text.color = Colors.GREEN
                self.load_available_filters()
                self.show_alert("✅ LOB regenerado exitosamente!")
            else:
                self.status_text.value = "❌ Error regenerando LOB"
                self.status_text.color = Colors.RED
                self.show_alert(f"❌ Error: {result.stderr[:200]}...")
                
        except Exception as ex:
            self.status_text.value = f"❌ Error: {str(ex)}"
            self.status_text.color = Colors.RED
            self.show_alert(f"❌ Error regenerando: {str(ex)}")
        
        finally:
            self.regenerate_button.disabled = False
            self.page.update()
    
    def export_filtered_data(self, e):
        """Exportar datos filtrados"""
        if self.filtered_items is None or len(self.filtered_items) == 0:
            self.show_alert("❌ No hay datos para exportar")
            return
        
        self.export_button.disabled = True
        self.status_text.value = "📊 Exportando..."
        self.status_text.color = Colors.ORANGE
        self.page.update()
        
        try:
            # Preparar datos
            export_data = self.filtered_items.copy()
            
            # Generar nombre
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filters_str = "_".join([f"{k}={v}" for k, v in self.current_filters.items() if v != "Todos"])
            if not filters_str:
                filters_str = "AllData"
            
            filename = f"LOB_Dashboard_{filters_str}_{timestamp}.csv"
            
            # Exportar
            export_data.to_csv(filename, index=False)
            
            # Estadísticas
            total_records = len(export_data)
            critical_count = len(export_data[export_data['status'] == 'critical'])
            total_value = export_data['inventory_value'].sum()
            
            self.status_text.value = f"✅ Exportado: {filename}"
            self.status_text.color = Colors.GREEN
            
            # Mostrar resumen
            summary_text = f"📊 Datos Exportados:\n\n"
            summary_text += f"📁 Archivo: {filename}\n"
            summary_text += f"📦 Registros: {total_records:,}\n"
            summary_text += f"🔴 Críticos: {critical_count:,}\n"
            summary_text += f"💰 Valor: ${total_value:,.2f}"
            
            self.show_alert(summary_text)
            
        except Exception as ex:
            self.status_text.value = f"❌ Error exportando"
            self.status_text.color = Colors.RED
            self.show_alert(f"❌ Error: {str(ex)}")
        
        finally:
            self.export_button.disabled = False
            self.page.update()
    
    def show_alert(self, message: str):
        """Mostrar alerta"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("💬 Información"),
            content=ft.Text(message),
            actions=[ft.TextButton("OK", on_click=close_dlg)]
        )
        
        self.page.dialog = dlg
        dlg.open = True
        self.page.update()

def main(page: ft.Page):
    """Función principal"""
    OptimizedLOBDashboard(page)

if __name__ == "__main__":
    ft.app(target=main, assets_dir="assets")