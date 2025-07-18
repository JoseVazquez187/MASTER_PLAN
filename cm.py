import flet as ft
import pandas as pd
import sqlite3
import datetime as dt
import os
from datetime import datetime, timedelta

class DataManager:
    """Gestor de datos desde R4Database"""
    
    def __init__(self):
        self.db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        self.df = None
        self.reason_mapping = {
            "RFC": "RETURN FOR CUSTOMER",
            "OEE": "ORDER ENTRY ERROR", 
            "INE": "INVOICE ERROR",
            "PNS": "PARTS NOT SHIPPED",
            "CBC": "CAUSED BY CUSTOMER",
            "TDL": "TRANSIT DAMAGE LOSS"
        }
    
    def load_data(self):
        """Carga y procesa datos desde R4Database - CORREGIDO"""
        try:
            print("🔄 Conectando a R4Database...")
            conn = sqlite3.connect(self.db_path)
            
            # Cargar Sales Orders - SOLO CON CANTIDAD ABIERTA > 0
            so_query = """
            SELECT SO_No, Ln, Cust_PO, Req_Dt, Spr_CD, ML, 
            Item_Number, Description, PlanType, Opn_Q, Issue_Q, OH_Netable
            FROM sales_order_table 
            WHERE Ord_Cd IN ('M20', 'M55') 
            AND Opn_Q > 0
            """
            so_df = pd.read_sql_query(so_query, conn)
            so_df['SO_Line'] = so_df['SO_No'].astype(str) + '-' + so_df['Ln'].astype(str)
            print(f"📊 Sales Orders con Opn_Q > 0: {len(so_df)}")
            
            # Cargar Credit Memos
            cm_query = """
            SELECT SO_No, Line, Invoice_No, CM_Reason, User_Id, Issue_Date, Invoice_Line_Memo
            FROM Credit_Memos
            WHERE Invoice_No IS NOT NULL AND Invoice_No != ''
            """
            cm_df = pd.read_sql_query(cm_query, conn)
            cm_df['SO_Line'] = cm_df['SO_No'].astype(str) + '-' + cm_df['Line'].astype(str)
            print(f"📋 Credit Memos: {len(cm_df)}")
            
            # Join datos - INNER JOIN para solo coincidencias
            merged_df = pd.merge(so_df, cm_df, on='SO_Line', how='inner')
            print(f"🔗 Coincidencias SO con Opn_Q > 0 y CM: {len(merged_df)}")
            
            # Procesar datos
            merged_df['CM_Reason'] = merged_df['CM_Reason'].str.upper()
            merged_df['CM_Reason_Full'] = merged_df['CM_Reason'].map(self.reason_mapping).fillna('CHECK REASON')
            merged_df['Issue_Date'] = pd.to_datetime(merged_df['Issue_Date'], errors='coerce')
            merged_df['Req_Dt'] = pd.to_datetime(merged_df['Req_Dt'], errors='coerce')
            
            # Agregar información adicional para análisis
            merged_df['Has_Open_Qty'] = merged_df['Opn_Q'] > 0
            merged_df['Value_Impact'] = merged_df['Issue_Q'] * 50  # Estimación de valor unitario
            
            conn.close()
            self.df = merged_df
            
            print(f"✅ Datos procesados finales: {len(merged_df)} registros")
            print(f"📊 Facturas únicas: {merged_df['Invoice_No'].nunique()}")
            print(f"📦 Total Items (Issue_Q): {merged_df['Issue_Q'].sum()}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error: {e}")
            return False
    
    def get_kpis(self):
        """Calcula KPIs principales - CORREGIDO"""
        if self.df is None or self.df.empty:
            return {
                'total_cms': 0,
                'total_items': 0,
                'unique_suppliers': 0,
                'top_reason': 'N/A',
                'avg_issue_qty': 0,
                'avg_days_to_today': 0
            }
        
        # Cálculo de días promedio hasta hoy
        today = pd.Timestamp.now()
        self.df['days_to_today'] = (today - self.df['Issue_Date']).dt.days
        avg_days = round(self.df['days_to_today'].mean(), 1)
        
        return {
            'total_cms': len(self.df),
            'total_items': int(self.df['Issue_Q'].sum()),
            'unique_suppliers': self.df['Spr_CD'].nunique(),
            'top_reason': self.df['CM_Reason_Full'].mode().iloc[0] if not self.df['CM_Reason_Full'].empty else 'N/A',
            'avg_issue_qty': round(self.df['Issue_Q'].mean(), 2),
            'avg_days_to_today': avg_days
        }
    
    def get_trends_data(self):
        """Obtiene datos para gráficos de tendencias"""
        if self.df is None or self.df.empty:
            return {
                'monthly': pd.Series(),
                'by_reason': pd.Series(),
                'by_supplier': pd.DataFrame()
            }
        
        # Por mes
        monthly = self.df.groupby(self.df['Issue_Date'].dt.to_period('M')).size()
        
        # Por razón
        by_reason = self.df['CM_Reason_Full'].value_counts().head(10)
        
        # Por supplier
        by_supplier = self.df.groupby('Spr_CD').agg({
            'Issue_Q': 'sum',
            'SO_Line': 'count'
        }).sort_values('SO_Line', ascending=False).head(15)
        
        return {
            'monthly': monthly,
            'by_reason': by_reason,
            'by_supplier': by_supplier
        }

class ExecutiveDashboard:
    """Dashboard Ejecutivo Completo"""
    
    def __init__(self, page: ft.Page):
        self.page = page
        self.data_manager = DataManager()
        self.setup_page()
        
    def setup_page(self):
        """Configura la página con tema dark y diseño moderno"""
        self.page.title = "🎯 Credit Memo Executive Dashboard"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.bgcolor = "#0a0a0a"
        self.page.window_width = 1600
        self.page.window_height = 1000
        self.page.window_maximized = True
        self.page.padding = 0
        
        # Estado
        self.status_text = ft.Text("🔄 Iniciando...", color="#00bcd4", size=16)
        self.kpi_container = ft.Container()
        self.charts_container = ft.Container()
        self.data_table_container = ft.Container()
        
        self.build_ui()
        self.load_dashboard_data()
    
    def build_ui(self):
        """Construye la interfaz moderna"""
        
        # Header con gradiente
        header = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.ANALYTICS, size=60, color="#00bcd4"),
                ft.Column([
                    ft.Text("CREDIT MEMO", size=36, weight=ft.FontWeight.BOLD, color="#ffffff"),
                    ft.Text("Master Plan Validation", size=18, color="#00bcd4")
                ], spacing=0),
                ft.Container(expand=True),
                ft.Row([
                    ft.Container(
                        content=ft.IconButton(
                            ft.Icons.REFRESH,
                            icon_color="#ffffff",
                            icon_size=32,
                            tooltip="Actualizar Datos",
                            on_click=self.refresh_data,
                            style=ft.ButtonStyle(
                                bgcolor="#00838f",
                                shape=ft.RoundedRectangleBorder(radius=12)
                            )
                        ),
                        width=60,
                        height=60
                    ),
                    ft.Container(width=10),
                    ft.Container(
                        content=ft.IconButton(
                            ft.Icons.DOWNLOAD,
                            icon_color="#ffffff",
                            icon_size=32,
                            tooltip="Descargar Reporte Ejecutivo",
                            on_click=self.export_executive_report,
                            style=ft.ButtonStyle(
                                bgcolor="#388e3c",
                                shape=ft.RoundedRectangleBorder(radius=12)
                            )
                        ),
                        width=60,
                        height=60
                    )
                ])
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=30,
            bgcolor="#1a1a1a",
            border=ft.border.only(bottom=ft.border.BorderSide(2, "#00bcd4"))
        )
        
        # Status bar
        status_bar = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.STORAGE, color="#4caf50"),
                self.status_text,
                ft.Container(expand=True),
                ft.Text(f"🕒 Última actualización: {datetime.now().strftime('%H:%M:%S')}", 
                    color="#757575", size=14)
            ]),
            padding=15,
            bgcolor="#1a1a1a"
        )
        
        # Tabs modernas
        tabs = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            indicator_color="#00bcd4",
            label_color="#00bcd4",
            unselected_label_color="#757575",
            tabs=[
                ft.Tab(
                    text="📊 OVERVIEW",
                    content=ft.Container(
                        content=ft.Column([
                            self.kpi_container,
                            self.charts_container
                        ], spacing=20, scroll=ft.ScrollMode.AUTO),
                        padding=20
                    )
                ),
                ft.Tab(
                    text="📋 DATA EXPLORER",
                    content=ft.Container(
                        content=self.data_table_container,
                        padding=20
                    )
                )
            ]
        )
        
        # Layout principal
        self.page.add(
            ft.Column([
                header,
                status_bar,
                ft.Container(content=tabs, expand=True)
            ], spacing=0)
        )
    
    def load_dashboard_data(self):
        """Carga datos y actualiza dashboard"""
        self.status_text.value = "🔄 Cargando datos desde R4Database..."
        self.page.update()
        
        if self.data_manager.load_data():
            self.status_text.value = "✅ Datos cargados exitosamente"
            self.update_kpis()
            self.update_charts()
            self.update_data_table()
        else:
            self.status_text.value = "❌ Error cargando datos"
        
        self.page.update()
    
    def update_kpis(self):
        """Actualiza KPIs con diseño moderno"""
        kpis = self.data_manager.get_kpis()
        
        # Solo 5 cards relevantes - SIN INVOICES
        kpi_cards = [
            self.create_kpi_card("💼 TOTAL CMs", f"{kpis['total_cms']:,}", "#00bcd4", "Credit Memos únicos"),
            self.create_kpi_card("📦 TOTAL ISSUE QTY", f"{kpis['total_items']:,}", "#ff9800", "Cantidad total emitida"),
            self.create_kpi_card("🏭 SUPPLIERS", f"{kpis['unique_suppliers']:,}", "#4caf50", "Proveedores afectados"),
            self.create_kpi_card("⚡ TOP REASON", kpis['top_reason'][:12], "#f44336", "Razón principal"),
            self.create_kpi_card("📅 AVG DAYS", f"{kpis['avg_days_to_today']}", "#9c27b0", "Días promedio a hoy")
        ]
        
        self.kpi_container.content = ft.Container(
            content=ft.Row(kpi_cards, spacing=20, wrap=True),
            margin=ft.margin.only(bottom=30)
        )
    
    def create_kpi_card(self, title, value, color, subtitle):
        """Crea tarjeta KPI moderna con glassmorphism"""
        return ft.Container(
            content=ft.Column([
                ft.Text(title, size=12, weight=ft.FontWeight.BOLD, color="#bdbdbd"),
                ft.Text(value, size=28, weight=ft.FontWeight.BOLD, color=color),
                ft.Text(subtitle, size=10, color="#757575")
            ], spacing=8, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            width=250,
            height=120,
            padding=20,
            bgcolor="#1a1a1a",
            border=ft.border.all(1, color),
            border_radius=15,
            shadow=ft.BoxShadow(
                spread_radius=1,
                blur_radius=15,
                color="#424242",
                offset=ft.Offset(0, 5)
            )
        )
    
    def update_charts(self):
        """Actualiza gráficos interactivos"""
        trends = self.data_manager.get_trends_data()
        
        # Gráfico de razones (Donut Chart)
        reason_chart = self.create_reason_donut_chart(trends['by_reason'])
        
        # Gráfico de tendencias temporales
        temporal_chart = self.create_temporal_chart(trends['monthly'])
        
        # Gráfico de suppliers
        supplier_chart = self.create_supplier_chart(trends['by_supplier'])
        
        charts_layout = ft.Column([
            ft.Row([reason_chart, temporal_chart], spacing=30),
            ft.Container(height=20),
            supplier_chart
        ])
        
        self.charts_container.content = charts_layout
    
    def create_reason_donut_chart(self, data):
        """Crea gráfico donut de razones"""
        if data.empty:
            return ft.Container(
                content=ft.Text("No hay datos de razones", color="#ffffff"),
                padding=25,
                bgcolor="#1a1a1a",
                border_radius=15
            )
        
        # Colores vibrantes
        colors = ['#00bcd4', '#ff5722', '#4caf50', '#ff9800', '#9c27b0', '#2196f3']
        
        chart_items = []
        total = data.sum()
        
        for i, (reason, count) in enumerate(data.head(6).items()):
            percentage = (count / total) * 100
            color = colors[i % len(colors)]
            
            chart_items.append(
                ft.Container(
                    content=ft.Row([
                        ft.Container(width=20, height=20, bgcolor=color, border_radius=10),
                        ft.Column([
                            ft.Text(reason[:20] + "..." if len(reason) > 20 else reason, 
                                   size=14, weight=ft.FontWeight.BOLD, color="#ffffff"),
                            ft.Text(f"{count:,} casos ({percentage:.1f}%)", 
                                   size=12, color="#bdbdbd")
                        ], spacing=2)
                    ], spacing=15),
                    padding=ft.padding.symmetric(vertical=12, horizontal=15),
                    margin=ft.margin.only(bottom=8),
                    bgcolor="#1a1a1a",
                    border_radius=10
                )
            )
        
        return ft.Container(
            content=ft.Column([
                ft.Text("📊 DISTRIBUCIÓN POR RAZONES", size=18, weight=ft.FontWeight.BOLD, color="#ffffff"),
                ft.Container(height=15),
                ft.Column(chart_items, scroll=ft.ScrollMode.AUTO, height=300)
            ]),
            width=500,
            padding=25,
            bgcolor="#1a1a1a",
            border=ft.border.all(1, "#00bcd4"),
            border_radius=15
        )
    
    def create_temporal_chart(self, data):
        """Crea gráfico de línea temporal"""
        if data.empty:
            return ft.Container(
                content=ft.Text("No hay datos temporales", color="#ffffff"),
                padding=25,
                bgcolor="#1a1a1a",
                border_radius=15
            )
        
        chart_items = []
        max_val = data.max() if not data.empty else 1
        
        for period, count in data.tail(12).items():
            bar_height = (count / max_val) * 200
            
            chart_items.append(
                ft.Container(
                    content=ft.Column([
                        ft.Container(
                            width=30,
                            height=bar_height,
                            bgcolor="#00bcd4",
                            border_radius=5
                        ),
                        ft.Text(str(period)[-7:], size=10, color="#bdbdbd")
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    margin=ft.margin.only(right=10)
                )
            )
        
        return ft.Container(
            content=ft.Column([
                ft.Text("📈 TENDENCIA TEMPORAL", size=18, weight=ft.FontWeight.BOLD, color="#ffffff"),
                ft.Container(height=15),
                ft.Row(chart_items, scroll=ft.ScrollMode.AUTO)
            ]),
            width=500,
            padding=25,
            bgcolor="#1a1a1a",
            border=ft.border.all(1, "#ff9800"),
            border_radius=15
        )
    
    def create_supplier_chart(self, data):
        """Crea tabla de suppliers con impacto visual"""
        if data.empty:
            return ft.Container(
                content=ft.Text("No hay datos de suppliers", color="#ffffff"),
                padding=25,
                bgcolor="#1a1a1a",
                border_radius=15
            )
        
        rows = []
        for supplier, row in data.head(10).iterrows():
            impact_color = "#f44336" if row['SO_Line'] > 10 else "#ff9800" if row['SO_Line'] > 5 else "#4caf50"
            impact_text = "🔴 ALTO" if row['SO_Line'] > 10 else "🟡 MEDIO" if row['SO_Line'] > 5 else "🟢 BAJO"
            
            rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(supplier, color="#ffffff", weight=ft.FontWeight.BOLD)),
                        ft.DataCell(ft.Text(f"{row['SO_Line']:,}", color="#00bcd4")),
                        ft.DataCell(ft.Text(f"{int(row['Issue_Q']):,}", color="#ff9800")),
                        ft.DataCell(
                            ft.Container(
                                content=ft.Text(impact_text, size=12, weight=ft.FontWeight.BOLD),
                                padding=8,
                                bgcolor="#2a2a2a",
                                border_radius=8
                            )
                        )
                    ]
                )
            )
        
        table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("🏭 SUPPLIER", weight=ft.FontWeight.BOLD, color="#ffffff")),
                ft.DataColumn(ft.Text("📋 CMs", weight=ft.FontWeight.BOLD, color="#ffffff")),
                ft.DataColumn(ft.Text("📦 ITEMS", weight=ft.FontWeight.BOLD, color="#ffffff")),
                ft.DataColumn(ft.Text("⚠️ IMPACTO", weight=ft.FontWeight.BOLD, color="#ffffff"))
            ],
            rows=rows,
            bgcolor="#1a1a1a",
            border=ft.border.all(1, "#9c27b0"),
            border_radius=10,
            heading_row_color="#2a2a2a"
        )
        
        return ft.Container(
            content=ft.Column([
                ft.Text("🏭 TOP SUPPLIERS BY IMPACT", size=18, weight=ft.FontWeight.BOLD, color="#ffffff"),
                ft.Container(height=15),
                table
            ]),
            padding=25,
            bgcolor="#1a1a1a",
            border=ft.border.all(1, "#9c27b0"),
            border_radius=15
        )
    
    def update_data_table(self):
        """Actualiza tabla de datos con scroll SIMPLE Y FUNCIONAL"""
        if self.data_manager.df is None or self.data_manager.df.empty:
            self.data_table_container.content = ft.Text("No hay datos disponibles", color="#ffffff")
            return
        
        # Campo de búsqueda
        search_field = ft.TextField(
            label="🔍 Buscar SO Line, Invoice, Supplier, User...",
            width=500,
            color="#ffffff",
            bgcolor="#2a2a2a",
            border_color="#00bcd4",
            on_change=self.search_table
        )
        
        # Crear todas las filas directamente
        self.create_all_table_rows()
        
        # Layout simple con scroll
        self.data_table_container.content = ft.Column([
            ft.Text("📋 CREDIT MEMOS DATA EXPLORER", size=20, weight=ft.FontWeight.BOLD, color="#ffffff"),
            ft.Text(f"📊 Total: {len(self.data_manager.df):,} registros", color="#bdbdbd", size=12),
            ft.Container(height=10),
            search_field,
            ft.Container(height=10),
            self.scrollable_table
        ])
    
    def create_all_table_rows(self):
        """Crea tabla con TODAS las filas y scroll FUNCIONAL"""
        df = self.data_manager.df
        
        rows = []
        for _, row in df.iterrows():
            # Colores
            opn_q_color = "#4caf50" if row['Opn_Q'] > 0 else "#f44336"
            days_diff = (pd.Timestamp.now() - row['Issue_Date']).days
            days_color = "#4caf50" if days_diff < 30 else "#ff9800" if days_diff < 90 else "#f44336"
            
            rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(str(row['SO_Line']), color="#00bcd4", size=11)),
                        ft.DataCell(ft.Text(str(row['Invoice_No']), color="#ffffff", size=11)),
                        ft.DataCell(ft.Text(str(row['CM_Reason_Full'])[:12], color="#ff9800", size=11)),
                        ft.DataCell(ft.Text(str(row['Spr_CD']), color="#4caf50", size=11)),
                        ft.DataCell(ft.Text(str(row['User_Id']), color="#9c27b0", size=11)),
                        ft.DataCell(ft.Text(f"{int(row['Issue_Q'])}", color="#ffeb3b", size=11)),
                        ft.DataCell(ft.Text(f"{int(row['Opn_Q'])}", color=opn_q_color, size=11)),
                        ft.DataCell(ft.Text(f"{days_diff}d", color=days_color, size=11)),
                        ft.DataCell(ft.Text(str(row['Issue_Date'].date()) if pd.notna(row['Issue_Date']) else 'N/A', color="#bdbdbd", size=11))
                    ]
                )
            )
        
        # Tabla simple
        data_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("SO Line", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Invoice", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Reason", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Supplier", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("User", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Issue", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Open", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Days", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Date", weight=ft.FontWeight.BOLD, color="#ffffff", size=11))
            ],
            rows=rows,
            bgcolor="#1a1a1a",
            border_radius=10,
            heading_row_color="#2a2a2a"
        )
        
        # SCROLL REAL - Envuelto en Column con scroll explícito
        self.scrollable_table = ft.Container(
            content=ft.Column(
                controls=[data_table],
                scroll=ft.ScrollMode.ALWAYS,  # ESTO es la clave
                auto_scroll=True
            ),
            height=400,
            bgcolor="#1a1a1a",
            border_radius=10,
            padding=10
        )
    
    def search_table(self, e):
        """Búsqueda simple y efectiva"""
        search_term = e.control.value.lower().strip()
        
        if not search_term:
            filtered_df = self.data_manager.df.copy()
        else:
            mask = (
                self.data_manager.df['SO_Line'].astype(str).str.lower().str.contains(search_term, na=False) |
                self.data_manager.df['Invoice_No'].astype(str).str.lower().str.contains(search_term, na=False) |
                self.data_manager.df['Spr_CD'].astype(str).str.lower().str.contains(search_term, na=False) |
                self.data_manager.df['User_Id'].astype(str).str.lower().str.contains(search_term, na=False)
            )
            filtered_df = self.data_manager.df[mask]
        
        # Recrear tabla con datos filtrados
        self.create_filtered_table(filtered_df)
        self.page.update()
    
    def create_filtered_table(self, df):
        """Crea tabla filtrada CON SCROLL"""
        rows = []
        for _, row in df.iterrows():
            opn_q_color = "#4caf50" if row['Opn_Q'] > 0 else "#f44336"
            days_diff = (pd.Timestamp.now() - row['Issue_Date']).days
            days_color = "#4caf50" if days_diff < 30 else "#ff9800" if days_diff < 90 else "#f44336"
            
            rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(str(row['SO_Line']), color="#00bcd4", size=11)),
                        ft.DataCell(ft.Text(str(row['Invoice_No']), color="#ffffff", size=11)),
                        ft.DataCell(ft.Text(str(row['CM_Reason_Full'])[:12], color="#ff9800", size=11)),
                        ft.DataCell(ft.Text(str(row['Spr_CD']), color="#4caf50", size=11)),
                        ft.DataCell(ft.Text(str(row['User_Id']), color="#9c27b0", size=11)),
                        ft.DataCell(ft.Text(f"{int(row['Issue_Q'])}", color="#ffeb3b", size=11)),
                        ft.DataCell(ft.Text(f"{int(row['Opn_Q'])}", color=opn_q_color, size=11)),
                        ft.DataCell(ft.Text(f"{days_diff}d", color=days_color, size=11)),
                        ft.DataCell(ft.Text(str(row['Issue_Date'].date()) if pd.notna(row['Issue_Date']) else 'N/A', color="#bdbdbd", size=11))
                    ]
                )
            )
        
        filtered_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("SO Line", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Invoice", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Reason", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Supplier", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("User", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Issue", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Open", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Days", weight=ft.FontWeight.BOLD, color="#ffffff", size=11)),
                ft.DataColumn(ft.Text("Date", weight=ft.FontWeight.BOLD, color="#ffffff", size=11))
            ],
            rows=rows,
            bgcolor="#1a1a1a",
            border_radius=10,
            heading_row_color="#2a2a2a"
        )
        
        # Aplicar el mismo scroll a la tabla filtrada
        self.scrollable_table.content = ft.Column(
            controls=[filtered_table],
            scroll=ft.ScrollMode.ALWAYS,
            auto_scroll=True
        )
    
    def refresh_data(self, e):
        """Refresca todos los datos"""
        self.load_dashboard_data()
        self.show_snackbar("🔄 Datos actualizados desde R4Database", "#00bcd4")
    
    def export_executive_report(self, e):
        """Exporta reporte ejecutivo completo - ARREGLADO SIN OPTIONS"""
        try:
            if self.data_manager.df is None or self.data_manager.df.empty:
                self.show_snackbar("⚠️ No hay datos para exportar", "#ff9800")
                return
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"EXECUTIVE_CREDIT_MEMO_REPORT_{timestamp}.xlsx"
            
            # Mostrar progreso
            self.show_snackbar("🔄 Generando reporte ejecutivo...", "#00bcd4")
            
            kpis = self.data_manager.get_kpis()
            trends = self.data_manager.get_trends_data()
            
            try:
                # MÉTODO SIMPLE - Solo engine openpyxl
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    
                    # Hoja 1: Resumen Ejecutivo
                    executive_summary = pd.DataFrame([
                        ['RESUMEN EJECUTIVO', ''],
                        ['Total Credit Memos', kpis['total_cms']],
                        ['Total Issue Quantity', kpis['total_items']],
                        ['Suppliers Únicos', kpis['unique_suppliers']],
                        ['Razón Principal', kpis['top_reason']],
                        ['Promedio Issue Qty', kpis['avg_issue_qty']],
                        ['Promedio Days to Today', kpis['avg_days_to_today']],
                        ['', ''],
                        ['INFORMACIÓN', ''],
                        ['Solo SOs con Open Qty > 0', ''],
                        ['Days = días desde Issue Date hasta hoy', ''],
                        ['Reporte generado', datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
                    ], columns=['Métrica', 'Valor'])
                    executive_summary.to_excel(writer, sheet_name='Executive_Summary', index=False)
                    
                    # Hoja 2: Datos completos - SIN conversión de fechas problemática
                    export_df = self.data_manager.df.copy()
                    # Convertir fechas a string simple
                    if 'Issue_Date' in export_df.columns:
                        export_df['Issue_Date_Str'] = export_df['Issue_Date'].dt.strftime('%Y-%m-%d')
                        export_df = export_df.drop('Issue_Date', axis=1)  # Remover columna datetime original
                    
                    if 'Req_Dt' in export_df.columns:
                        export_df['Req_Dt_Str'] = export_df['Req_Dt'].dt.strftime('%Y-%m-%d')
                        export_df = export_df.drop('Req_Dt', axis=1)  # Remover columna datetime original
                    
                    export_df.to_excel(writer, sheet_name='Credit_Memos_Data', index=False)
                    
                    # Hoja 3: Análisis por razones
                    if not trends['by_reason'].empty:
                        reason_data = []
                        for reason, count in trends['by_reason'].items():
                            reason_data.append([reason, count])
                        
                        reason_df = pd.DataFrame(reason_data, columns=['CM_Reason', 'Count'])
                        reason_df.to_excel(writer, sheet_name='Analysis_by_Reason', index=False)
                    
                    # Hoja 4: Análisis por supplier
                    if not trends['by_supplier'].empty:
                        supplier_data = []
                        for supplier in trends['by_supplier'].index:
                            row = trends['by_supplier'].loc[supplier]
                            supplier_data.append([supplier, int(row['Issue_Q']), int(row['SO_Line'])])
                        
                        supplier_df = pd.DataFrame(supplier_data, columns=['Supplier', 'Total_Issue_Q', 'Total_CMs'])
                        supplier_df.to_excel(writer, sheet_name='Analysis_by_Supplier', index=False)
                    
                    # Hoja 5: Análisis por factura (método simple)
                    try:
                        invoice_data = []
                        for invoice in self.data_manager.df['Invoice_No'].unique():
                            invoice_subset = self.data_manager.df[self.data_manager.df['Invoice_No'] == invoice]
                            lines_count = len(invoice_subset)
                            total_issue_q = int(invoice_subset['Issue_Q'].sum())
                            supplier = invoice_subset['Spr_CD'].iloc[0]
                            issue_date = invoice_subset['Issue_Date'].iloc[0].strftime('%Y-%m-%d')
                            
                            invoice_data.append([invoice, lines_count, total_issue_q, supplier, issue_date])
                        
                        invoice_df = pd.DataFrame(invoice_data, columns=['Invoice_No', 'Lines_Count', 'Total_Issue_Q', 'Supplier', 'Issue_Date'])
                        invoice_df.to_excel(writer, sheet_name='Analysis_by_Invoice', index=False)
                    except Exception as invoice_error:
                        print(f"Warning: No se pudo crear análisis por factura: {invoice_error}")
                
                # Verificar archivo
                if os.path.exists(filename):
                    file_size = os.path.getsize(filename) / 1024  # KB
                    self.show_snackbar(f"✅ Reporte exportado: {filename} ({file_size:.1f} KB)", "#4caf50")
                    print(f"📁 Archivo guardado exitosamente: {filename}")
                else:
                    self.show_snackbar("❌ Error: Archivo no se creó", "#f44336")
                
            except PermissionError:
                self.show_snackbar("❌ Cierra el archivo Excel si está abierto e intenta de nuevo", "#f44336")
            except Exception as export_error:
                print(f"Error exportación: {export_error}")
                
                # RESPALDO: Exportación súper simple
                try:
                    simple_filename = f"SIMPLE_CM_REPORT_{timestamp}.xlsx"
                    simple_df = self.data_manager.df.copy()
                    
                    # Preparar datos básicos sin fechas datetime
                    simple_df['Issue_Date'] = simple_df['Issue_Date'].dt.strftime('%Y-%m-%d')
                    if 'Req_Dt' in simple_df.columns:
                        simple_df['Req_Dt'] = simple_df['Req_Dt'].dt.strftime('%Y-%m-%d')
                    
                    simple_df.to_excel(simple_filename, index=False, engine='openpyxl')
                    self.show_snackbar(f"✅ Exportación simple exitosa: {simple_filename}", "#4caf50")
                    print(f"📁 Archivo simple creado: {simple_filename}")
                    
                except Exception as simple_error:
                    self.show_snackbar(f"❌ Error total: {str(simple_error)}", "#f44336")
                    print(f"Error completo: {simple_error}")
            
        except Exception as e:
            self.show_snackbar(f"❌ Error general: {str(e)}", "#f44336")
            print(f"Error general en export: {e}")
    
    def show_snackbar(self, message, color):
        """Muestra mensaje de estado"""
        self.page.snack_bar = ft.SnackBar(
            content=ft.Text(message, color="#ffffff"),
            bgcolor=color
        )
        self.page.snack_bar.open = True
        self.page.update()

def main(page: ft.Page):
    """Función principal"""
    dashboard = ExecutiveDashboard(page)

if __name__ == "__main__":
    print("🚀 INICIANDO EXECUTIVE DASHBOARD")
    print("🎯 Modo Dark Theme Activado")
    print("📊 Conectando a R4Database...")
    ft.app(target=main, port=8084)