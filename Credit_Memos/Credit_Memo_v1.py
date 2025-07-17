import flet as ft
import pandas as pd
import sqlite3
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import json
from typing import Dict, List, Optional

class CreditMemoProcessor:
    """Clase para procesar los datos de credit memos con mejoras"""
    
    def __init__(self):
        self.reason_mapping = {
            "RFC": "RETURN FOR CUSTOMER",
            "OEE": "ORDER ENTRY ERROR", 
            "INE": "INVOICE ERROR",
            "PNS": "PARTS NOT SHIPPED",
            "CBC": "CAUSED BY CUSTOMER",
            "TDL": "TRANSIT DAMAGE LOSS"
        }
    
    def process_credit_memos(self, so_df: pd.DataFrame, credit_memo_df: pd.DataFrame) -> pd.DataFrame:
        """Procesa los credit memos con el código mejorado"""
        try:
            # Crear identificador único SO_Line
            credit_memo_df["SO-No"] = credit_memo_df["SO-No"].astype(str)
            credit_memo_df["Line"] = credit_memo_df["Line"].astype(str)
            credit_memo_df["SO_Line"] = credit_memo_df["SO-No"] + "-" + credit_memo_df["Line"]
            
            # Merge con sales orders
            credit_memos = so_df.merge(credit_memo_df, left_on="SO_Line", right_on="SO_Line", how="left")
            
            # Seleccionar columnas necesarias
            columns_to_keep = [
                "PO_Line", "SO_Line", "Req-Dt", "Spr-CD", "ML",
                "Item-Number", "Description_x", "PlanType", "Opn-Q", "Issue-Q",
                "OH Netable", "Invoice-No", "CM-Reason", "User-Id", "Issue-Date", "Invoice Line Memo"
            ]
            
            credit_memos = credit_memos[columns_to_keep]
            
            # Limpiar y mapear razones
            credit_memos["CM-Reason"] = credit_memos["CM-Reason"].str.upper()
            credit_memos["CM-Reason"] = credit_memos["CM-Reason"].map(self.reason_mapping).fillna("CHECK REASON")
            
            # Filtrar registros con invoice
            credit_memos["Invoice-No"] = credit_memos["Invoice-No"].fillna("")
            credit_memos = credit_memos.loc[credit_memos["Invoice-No"] != ""]
            
            # Convertir fechas
            credit_memos["Issue-Date"] = pd.to_datetime(credit_memos["Issue-Date"]).dt.date
            credit_memos["Req-Dt"] = pd.to_datetime(credit_memos["Req-Dt"], errors='coerce').dt.date
            
            return credit_memos
            
        except Exception as e:
            print(f"Error processing credit memos: {e}")
            return pd.DataFrame()

class DatabaseManager:
    """Gestor de conexión a base de datos"""
    
    def __init__(self, db_path: str = "R4Database.db"):
        self.db_path = db_path
    
    def get_connection(self):
        """Obtiene conexión a la base de datos"""
        return sqlite3.connect(self.db_path)
    
    def fetch_credit_memos(self) -> pd.DataFrame:
        """Obtiene datos de credit memos desde la base de datos"""
        try:
            with self.get_connection() as conn:
                query = """
                SELECT 
                    cm.*,
                    so.PO_Line,
                    so.Req_Dt as "Req-Dt",
                    so.Spr_CD as "Spr-CD",
                    so.ML,
                    so.Item_Number as "Item-Number",
                    so.Description as "Description_x",
                    so.PlanType,
                    so.Opn_Q as "Opn-Q",
                    so.Issue_Q as "Issue-Q",
                    so.OH_Netable as "OH Netable"
                FROM credit_memos cm
                LEFT JOIN sales_orders so ON cm.SO_Line = so.SO_Line
                WHERE cm.Invoice_No IS NOT NULL AND cm.Invoice_No != ''
                ORDER BY cm.Issue_Date DESC
                """
                return pd.read_sql_query(query, conn)
        except Exception as e:
            print(f"Error fetching data: {e}")
            return self._get_sample_data()
    
    def _get_sample_data(self) -> pd.DataFrame:
        """Datos de muestra para testing"""
        return pd.DataFrame({
            'PO_Line': ['PO001-1', 'PO002-1', 'PO003-1', 'PO004-1', 'PO005-1'],
            'SO_Line': ['SO001-1', 'SO002-1', 'SO003-1', 'SO004-1', 'SO005-1'],
            'Req-Dt': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19'],
            'Spr-CD': ['SPR001', 'SPR002', 'SPR003', 'SPR004', 'SPR005'],
            'ML': ['ML001', 'ML002', 'ML003', 'ML004', 'ML005'],
            'Item-Number': ['ITM001', 'ITM002', 'ITM003', 'ITM004', 'ITM005'],
            'Description_x': ['Item 1', 'Item 2', 'Item 3', 'Item 4', 'Item 5'],
            'PlanType': ['A', 'B', 'A', 'C', 'B'],
            'Opn-Q': [10, 20, 15, 30, 25],
            'Issue-Q': [5, 10, 8, 15, 12],
            'OH Netable': [5, 10, 7, 15, 13],
            'Invoice-No': ['INV001', 'INV002', 'INV003', 'INV004', 'INV005'],
            'CM-Reason': ['RFC', 'OEE', 'INE', 'PNS', 'CBC'],
            'User-Id': ['USER1', 'USER2', 'USER3', 'USER4', 'USER5'],
            'Issue-Date': ['2024-01-20', '2024-01-21', '2024-01-22', 'M024-01-23', '2024-01-24'],
            'Invoice Line Memo': ['Memo 1', 'Memo 2', 'Memo 3', 'Memo 4', 'Memo 5']
        })

class CreditMemoDashboard:
    """Dashboard principal para Credit Memos"""
    
    def __init__(self, page: ft.Page):
        self.page = page
        self.db_manager = DatabaseManager()
        self.processor = CreditMemoProcessor()
        self.df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()
        
        # Configurar página
        self.page.title = "Credit Memo Dashboard"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.padding = 20
        
        # Controles de filtros
        self.date_from = ft.DatePicker(
            first_date=datetime(2020, 1, 1),
            last_date=datetime(2030, 12, 31),
            on_change=self.on_filter_change
        )
        
        self.date_to = ft.DatePicker(
            first_date=datetime(2020, 1, 1),
            last_date=datetime(2030, 12, 31),
            on_change=self.on_filter_change
        )
        
        self.reason_filter = ft.Dropdown(
            label="Razón CM",
            width=200,
            on_change=self.on_filter_change
        )
        
        self.user_filter = ft.Dropdown(
            label="Usuario",
            width=200,
            on_change=self.on_filter_change
        )
        
        # Contenedores para gráficos
        self.metrics_container = ft.Container()
        self.chart_container = ft.Container()
        self.table_container = ft.Container()
        
        self.setup_page()
        self.load_data()
    
    def setup_page(self):
        """Configura la estructura de la página"""
        # Header
        header = ft.Container(
            content=ft.Row([
                ft.Icon(ft.icons.ASSESSMENT, size=40, color=ft.colors.BLUE_600),
                ft.Text("Credit Memo Dashboard", size=32, weight=ft.FontWeight.BOLD),
                ft.Container(expand=True),
                ft.ElevatedButton(
                    "Actualizar Datos",
                    icon=ft.icons.REFRESH,
                    on_click=self.refresh_data
                )
            ]),
            padding=20,
            bgcolor=ft.colors.SURFACE_VARIANT,
            border_radius=10,
            margin=ft.margin.only(bottom=20)
        )
        
        # Filtros
        filters = ft.Container(
            content=ft.Row([
                ft.ElevatedButton(
                    "Fecha Inicio",
                    icon=ft.icons.DATE_RANGE,
                    on_click=lambda _: self.date_from.pick_date()
                ),
                ft.ElevatedButton(
                    "Fecha Fin",
                    icon=ft.icons.DATE_RANGE,
                    on_click=lambda _: self.date_to.pick_date()
                ),
                self.reason_filter,
                self.user_filter,
                ft.ElevatedButton(
                    "Limpiar Filtros",
                    icon=ft.icons.CLEAR,
                    on_click=self.clear_filters
                )
            ], wrap=True, spacing=10),
            padding=20,
            bgcolor=ft.colors.SURFACE_VARIANT,
            border_radius=10,
            margin=ft.margin.only(bottom=20)
        )
        
        # Pestañas
        tabs = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=[
                ft.Tab(
                    text="Resumen",
                    icon=ft.icons.DASHBOARD,
                    content=ft.Column([
                        self.metrics_container,
                        self.chart_container
                    ])
                ),
                ft.Tab(
                    text="Datos Detallados",
                    icon=ft.icons.TABLE_CHART,
                    content=self.table_container
                )
            ]
        )
        
        # Agregar componentes a la página
        self.page.add(
            header,
            filters,
            self.date_from,
            self.date_to,
            tabs
        )
    
    def load_data(self):
        """Carga los datos desde la base de datos"""
        try:
            self.df = self.db_manager.fetch_credit_memos()
            
            # Procesar datos si es necesario
            if not self.df.empty:
                # Convertir fechas
                self.df['Issue-Date'] = pd.to_datetime(self.df['Issue-Date'], errors='coerce')
                self.df['Req-Dt'] = pd.to_datetime(self.df['Req-Dt'], errors='coerce')
                
                # Mapear razones
                self.df['CM-Reason'] = self.df['CM-Reason'].map(self.processor.reason_mapping).fillna('CHECK REASON')
            
            self.filtered_df = self.df.copy()
            self.update_filters()
            self.update_dashboard()
            
        except Exception as e:
            self.show_error(f"Error cargando datos: {e}")
    
    def update_filters(self):
        """Actualiza las opciones de los filtros"""
        if not self.df.empty:
            # Razones únicas
            reasons = ['Todas'] + sorted(self.df['CM-Reason'].dropna().unique().tolist())
            self.reason_filter.options = [ft.dropdown.Option(r) for r in reasons]
            
            # Usuarios únicos
            users = ['Todos'] + sorted(self.df['User-Id'].dropna().unique().tolist())
            self.user_filter.options = [ft.dropdown.Option(u) for u in users]
    
    def on_filter_change(self, e):
        """Maneja cambios en los filtros"""
        self.apply_filters()
        self.update_dashboard()
    
    def apply_filters(self):
        """Aplica los filtros seleccionados"""
        self.filtered_df = self.df.copy()
        
        # Filtro por fecha
        if self.date_from.value:
            self.filtered_df = self.filtered_df[
                self.filtered_df['Issue-Date'] >= pd.to_datetime(self.date_from.value)
            ]
        
        if self.date_to.value:
            self.filtered_df = self.filtered_df[
                self.filtered_df['Issue-Date'] <= pd.to_datetime(self.date_to.value)
            ]
        
        # Filtro por razón
        if self.reason_filter.value and self.reason_filter.value != 'Todas':
            self.filtered_df = self.filtered_df[
                self.filtered_df['CM-Reason'] == self.reason_filter.value
            ]
        
        # Filtro por usuario
        if self.user_filter.value and self.user_filter.value != 'Todos':
            self.filtered_df = self.filtered_df[
                self.filtered_df['User-Id'] == self.user_filter.value
            ]
    
    def update_dashboard(self):
        """Actualiza todos los componentes del dashboard"""
        self.update_metrics()
        self.update_charts()
        self.update_table()
        self.page.update()
    
    def update_metrics(self):
        """Actualiza las métricas principales"""
        if self.filtered_df.empty:
            return
        
        total_cms = len(self.filtered_df)
        total_items = self.filtered_df['Issue-Q'].sum()
        unique_invoices = self.filtered_df['Invoice-No'].nunique()
        
        metrics = ft.Row([
            self.create_metric_card("Total CMs", str(total_cms), ft.icons.RECEIPT, ft.colors.BLUE_600),
            self.create_metric_card("Total Items", str(int(total_items)), ft.icons.INVENTORY, ft.colors.GREEN_600),
            self.create_metric_card("Invoices", str(unique_invoices), ft.icons.DESCRIPTION, ft.colors.ORANGE_600),
        ], spacing=20)
        
        self.metrics_container.content = metrics
    
    def create_metric_card(self, title: str, value: str, icon, color):
        """Crea una tarjeta de métrica"""
        return ft.Container(
            content=ft.Column([
                ft.Icon(icon, size=40, color=color),
                ft.Text(value, size=24, weight=ft.FontWeight.BOLD),
                ft.Text(title, size=14, color=ft.colors.ON_SURFACE_VARIANT)
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            padding=20,
            bgcolor=ft.colors.SURFACE_VARIANT,
            border_radius=10,
            width=200,
            height=120
        )
    
    def update_charts(self):
        """Actualiza los gráficos"""
        if self.filtered_df.empty:
            return
        
        # Gráfico de razones
        reason_counts = self.filtered_df['CM-Reason'].value_counts()
        
        chart_row = ft.Row([
            ft.Container(
                content=ft.Column([
                    ft.Text("Distribución por Razón", size=16, weight=ft.FontWeight.BOLD),
                    ft.Container(
                        content=self.create_pie_chart(reason_counts),
                        height=300
                    )
                ]),
                bgcolor=ft.colors.SURFACE_VARIANT,
                border_radius=10,
                padding=20,
                expand=True
            ),
            ft.Container(
                content=ft.Column([
                    ft.Text("Tendencia por Fecha", size=16, weight=ft.FontWeight.BOLD),
                    ft.Container(
                        content=self.create_time_series(),
                        height=300
                    )
                ]),
                bgcolor=ft.colors.SURFACE_VARIANT,
                border_radius=10,
                padding=20,
                expand=True
            )
        ], spacing=20)
        
        self.chart_container.content = chart_row
    
    def create_pie_chart(self, data):
        """Crea un gráfico de pastel con los datos"""
        # Simular gráfico con containers coloreados
        colors = [ft.colors.BLUE_400, ft.colors.RED_400, ft.colors.GREEN_400, 
                 ft.colors.YELLOW_600, ft.colors.PURPLE_400, ft.colors.ORANGE_400]
        
        items = []
        for i, (reason, count) in enumerate(data.head(6).items()):
            color = colors[i % len(colors)]
            percentage = (count / data.sum()) * 100
            
            items.append(
                ft.Container(
                    content=ft.Row([
                        ft.Container(width=20, height=20, bgcolor=color, border_radius=5),
                        ft.Text(f"{reason}: {count} ({percentage:.1f}%)", size=12)
                    ]),
                    margin=5
                )
            )
        
        return ft.Column(items)
    
    def create_time_series(self):
        """Crea gráfico de serie temporal"""
        daily_counts = self.filtered_df.groupby(
            self.filtered_df['Issue-Date'].dt.date
        ).size().reset_index(name='count')
        
        items = []
        for _, row in daily_counts.head(10).iterrows():
            items.append(
                ft.Container(
                    content=ft.Row([
                        ft.Text(str(row['Issue-Date']), size=12, width=100),
                        ft.Container(
                            width=row['count'] * 10,
                            height=20,
                            bgcolor=ft.colors.BLUE_400,
                            border_radius=5
                        ),
                        ft.Text(str(row['count']), size=12)
                    ]),
                    margin=2
                )
            )
        
        return ft.Column(items, scroll=ft.ScrollMode.AUTO)
    
    def update_table(self):
        """Actualiza la tabla de datos"""
        if self.filtered_df.empty:
            self.table_container.content = ft.Text("No hay datos disponibles")
            return
        
        # Crear tabla con paginación
        rows = []
        for _, row in self.filtered_df.head(50).iterrows():
            rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(str(row['SO_Line']))),
                        ft.DataCell(ft.Text(str(row['Invoice-No']))),
                        ft.DataCell(ft.Text(str(row['CM-Reason']))),
                        ft.DataCell(ft.Text(str(row['User-Id']))),
                        ft.DataCell(ft.Text(str(row['Issue-Date'].date() if pd.notna(row['Issue-Date']) else ''))),
                        ft.DataCell(ft.Text(str(row['Issue-Q']))),
                    ]
                )
            )
        
        data_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("SO Line")),
                ft.DataColumn(ft.Text("Invoice No")),
                ft.DataColumn(ft.Text("CM Reason")),
                ft.DataColumn(ft.Text("User ID")),
                ft.DataColumn(ft.Text("Issue Date")),
                ft.DataColumn(ft.Text("Issue Qty")),
            ],
            rows=rows,
            border=ft.border.all(1, ft.colors.OUTLINE),
            border_radius=10,
            vertical_lines=ft.border.BorderSide(1, ft.colors.OUTLINE),
            horizontal_lines=ft.border.BorderSide(1, ft.colors.OUTLINE),
        )
        
        self.table_container.content = ft.Column([
            ft.Text(f"Mostrando {len(rows)} de {len(self.filtered_df)} registros", 
                size=14, color=ft.colors.ON_SURFACE_VARIANT),
            ft.Container(
                content=data_table,
                height=400,
                padding=10,
                bgcolor=ft.colors.SURFACE_VARIANT,
                border_radius=10
            )
        ])
    
    def clear_filters(self, e):
        """Limpia todos los filtros"""
        self.reason_filter.value = None
        self.user_filter.value = None
        self.date_from.value = None
        self.date_to.value = None
        self.filtered_df = self.df.copy()
        self.update_dashboard()
    
    def refresh_data(self, e):
        """Refresca los datos"""
        self.load_data()
    
    def show_error(self, message: str):
        """Muestra un mensaje de error"""
        self.page.snack_bar = ft.SnackBar(
            content=ft.Text(message),
            bgcolor=ft.colors.ERROR,
        )
        self.page.snack_bar.open = True
        self.page.update()

def main(page: ft.Page):
    """Función principal de la aplicación"""
    dashboard = CreditMemoDashboard(page)

if __name__ == "__main__":
    ft.app(target=main, port=8000)