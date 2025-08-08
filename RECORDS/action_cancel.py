import flet as ft
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import numpy as np
import os
import tempfile
import webbrowser

# === Configuración del tema gerencial dark ===
COLORS = {
    'primary': '#1e293b',      # Slate 800
    'secondary': '#334155',    # Slate 700
    'accent': '#0ea5e9',       # Sky 500
    'success': '#10b981',      # Emerald 500
    'warning': '#f59e0b',      # Amber 500
    'error': '#ef4444',        # Red 500
    'surface': '#0f172a',      # Slate 900
    'text_primary': '#f8fafc', # Slate 50
    'text_secondary': '#94a3b8' # Slate 400
}

class WOCancellationDashboard:
    def __init__(self):
        # self.db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        self.db_path = r"C:\Users\J.Vazquez\Desktop\R4Database.db"
        self.df_original = None
        self.df_enriched = None

    def load_and_process_data(self):
        try:
            conn = sqlite3.connect(self.db_path)
            query = """
            SELECT
                Días_Sin_Ejecutar,
                Fecha_Primera_Aparicion,
                Entity_WorkOrder,
                REF as WO_Number,
                "ITEM-NO" as Item_Number,
                "ITEM-DESC" as Item_Description,
                PlanTp,
                "Std-Cost" as Std_Cost,
                "Act-Qty" as Act_Qty,
                "Fm-Qty" as Fm_Qty,
                "To-Qty" as To_Qty,
                OH,
                ABCD,
                MLI,
                Responsible_Final,
                Email_Final,
                Total_ValQtyIssued
            FROM Cancelation_Information_WO
            ORDER BY Días_Sin_Ejecutar DESC, Fecha_Primera_Aparicion DESC
            """
            self.df_original = pd.read_sql_query(query, conn)
            conn.close()
            if len(self.df_original) == 0:
                print("⚠️ No se encontraron registros en Cancelation_Information_WO")
                return False
            self.process_enriched_data()
            print(f"✅ Datos procesados: {len(self.df_original)} WOs con acciones de cancelación")
            return True
        except Exception as e:
            print(f"❌ Error cargando datos de cancelaciones: {e}")
            return False

    def process_enriched_data(self):
        df = self.df_original.copy()
        df['Fecha_Primera_Aparicion'] = pd.to_datetime(df['Fecha_Primera_Aparicion'], errors='coerce')
        numeric_columns = ['Días_Sin_Ejecutar', 'Std_Cost', 'Act_Qty', 'Fm_Qty', 'To_Qty', 'Total_ValQtyIssued']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df['Action_Value'] = df['Act_Qty'] * df['Std_Cost']
        df['Total_Potential_Value'] = df['Action_Value'] + df['Total_ValQtyIssued']
        df['Aging_Category'] = df['Días_Sin_Ejecutar'].apply(self._categorize_aging)
        df['Value_Category'] = df['Action_Value'].apply(self._categorize_value)
        df['Plan_Category'] = df['PlanTp'].apply(self._categorize_plan_type)
        df['Responsible_Status'] = df.apply(self._determine_responsible_status, axis=1)
        df['Criticality_Level'] = df.apply(self._determine_criticality, axis=1)
        df['Responsible_Final'] = df['Responsible_Final'].fillna('Sin Asignar')
        df['Email_Final'] = df['Email_Final'].fillna('No Email')
        self.df_enriched = df

    def _categorize_aging(self, days):
        if pd.isna(days) or days == 0:
            return 'Reciente'
        if days <= 7:
            return '1 Semana'
        elif days <= 30:
            return '1 Mes'
        elif days <= 90:
            return '3 Meses'
        elif days <= 180:
            return '6 Meses'
        else:
            return 'Crítico (>6M)'
    def _categorize_value(self, value):
        if pd.isna(value) or value == 0:
            return 'Sin Valor'
        if value < 100:
            return 'Bajo (<$100)'
        elif value < 1000:
            return 'Medio ($100-$1K)'
        elif value < 10000:
            return 'Alto ($1K-$10K)'
        else:
            return 'Crítico (>$10K)'
    def _categorize_plan_type(self, plan_type):
        if pd.isna(plan_type) or plan_type == '':
            return 'Sin Tipo'
        plan_str = str(plan_type).upper()
        if 'PM' in plan_str or 'PREVENTIVE' in plan_str:
            return 'Preventivo'
        elif 'CM' in plan_str or 'CORRECTIVE' in plan_str:
            return 'Correctivo'
        elif 'EMERGENCY' in plan_str or 'EMERG' in plan_str:
            return 'Emergencia'
        else:
            return 'Otro'
    def _determine_responsible_status(self, row):
        if pd.notna(row['Responsible_Final']) and row['Responsible_Final'] != 'Sin Asignar':
            if pd.notna(row['Email_Final']) and row['Email_Final'] != 'No Email':
                return 'Completo'
            else:
                return 'Sin Email'
        else:
            return 'Sin Asignar'
    def _determine_criticality(self, row):
        days = row['Días_Sin_Ejecutar']
        value = row['Action_Value']
        if pd.isna(days): days = 0
        if pd.isna(value): value = 0
        score = 0
        if days > 180: score += 4
        elif days > 90: score += 3
        elif days > 30: score += 2
        elif days > 7: score += 1
        if value > 10000: score += 3
        elif value > 1000: score += 2
        elif value > 100: score += 1
        if score >= 6: return 'Crítico'
        elif score >= 4: return 'Alto'
        elif score >= 2: return 'Medio'
        else: return 'Bajo'
    def get_user_metrics(self):
        if self.df_enriched is None or len(self.df_enriched) == 0:
            return pd.DataFrame()
        df = self.df_enriched
        df_with_users = df[
            (df['Responsible_Final'].notna()) & 
            (df['Responsible_Final'] != 'Sin Asignar') & 
            (df['Responsible_Final'] != '')
        ].copy()
        if len(df_with_users) == 0:
            return pd.DataFrame()
        user_metrics = df_with_users.groupby(['Responsible_Final', 'Email_Final']).agg({
            'WO_Number': ['count', 'nunique'],
            'Action_Value': ['sum', 'mean'],
            'Total_Potential_Value': 'sum',
            'Días_Sin_Ejecutar': ['mean', 'max'],
            'Criticality_Level': [
                lambda x: (x == 'Crítico').sum(),
                lambda x: (x == 'Alto').sum(),
                lambda x: (x == 'Medio').sum()
            ],
            'Aging_Category': lambda x: (x == 'Crítico (>6M)').sum(),
            'Entity_WorkOrder': 'nunique'
        }).reset_index()
        user_metrics.columns = [
            'Responsible', 'Email', 'Total_Actions', 'Unique_WOs', 'Total_Value', 'Avg_Value',
            'Total_Potential', 'Avg_Days', 'Max_Days', 'Critical_Actions',
            'High_Actions', 'Medium_Actions', 'Critical_Aging', 'Unique_Entities'
        ]
        user_metrics['Risk_Score'] = (
            (user_metrics['Critical_Actions'] * 5) + 
            (user_metrics['High_Actions'] * 3) + 
            (user_metrics['Critical_Aging'] * 4) +
            (user_metrics['Avg_Days'] / 30)
        ).round(1)
        user_metrics['Critical_Rate'] = (
            (user_metrics['Critical_Actions'] / user_metrics['Total_Actions']) * 100
        ).round(1)
        user_metrics['Avg_Value_Per_Action'] = (
            user_metrics['Total_Value'] / user_metrics['Total_Actions']
        ).round(0)
        user_metrics['Efficiency_Score'] = (
            100 - ((user_metrics['Avg_Days'] / user_metrics['Max_Days'].max()) * 100)
        ).round(1)
        user_metrics = user_metrics.sort_values('Risk_Score', ascending=False).reset_index(drop=True)
        user_metrics['Ranking'] = range(1, len(user_metrics) + 1)
        return user_metrics

def create_compact_metric_card(title, value, subtitle="", color=COLORS['accent'], icon="📊"):
    return ft.Card(
        content=ft.Container(
            padding=ft.padding.all(15),
            bgcolor=COLORS['primary'],
            border_radius=12,
            content=ft.Column([
                ft.Row([
                    ft.Text(icon, size=20),
                    ft.Text(title, size=12, color=COLORS['text_secondary'], weight=ft.FontWeight.W_500),
                ], spacing=8),
                ft.Text(str(value), size=24, color=color, weight=ft.FontWeight.BOLD),
                ft.Text(subtitle, size=10, color=COLORS['text_secondary']) if subtitle else ft.Container()
            ], spacing=5, tight=True)
        ),
        elevation=8,
        width=180,
        height=120
    )

def main(page: ft.Page):
    page.title = "WO Cancellation Actions Dashboard"
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = COLORS['surface']
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20

    dashboard = WOCancellationDashboard()
    dashboard.load_and_process_data()  # Carga datos al arrancar

    def create_user_metrics_tab():
        user_container = ft.Column(spacing=20)
        user_metrics_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Rank", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Responsable", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Email", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Total Acciones", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("WOs Únicos", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Valor Total", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Críticas", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Promedio Días", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Máx Días", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Risk Score", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Tasa Crítica %", color=COLORS['text_primary']))
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=8,
        )

        def descargar_excel(e=None):
            user_metrics = dashboard.get_user_metrics()
            if len(user_metrics) == 0:
                page.snack_bar = ft.SnackBar(ft.Text("No hay datos para exportar"), bgcolor=COLORS['error'])
                page.snack_bar.open = True
                page.update()
                return
            # Archivo temporal en escritorio
            temp_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            file_path = os.path.join(temp_dir, "WOCancel_Responsables.xlsx")
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                user_metrics.to_excel(writer, index=False, sheet_name="Resumen")
                # Verifica si hay información detallada
                if dashboard.df_enriched is not None and len(dashboard.df_enriched) > 0:
                    dashboard.df_enriched.to_excel(writer, index=False, sheet_name="Detalle")
            # Abre el archivo con el programa predeterminado (en Windows, Excel)
            webbrowser.open(f'file://{file_path}')
            page.snack_bar = ft.SnackBar(ft.Text(f"Archivo exportado: {file_path}"), bgcolor=COLORS['success'])
            page.snack_bar.open = True
            page.update()




        def update_user_metrics():
            user_container.controls.clear()
            user_metrics_table.rows.clear()
            user_metrics = dashboard.get_user_metrics()
            if len(user_metrics) == 0:
                user_container.controls.append(
                    ft.Container(
                        content=ft.Text("⚠️ No hay responsables asignados a acciones de cancelación",
                                        size=18, color=COLORS['warning']),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=8
                    )
                )
                page.update()
                return
            total_responsibles = len(user_metrics)
            responsibles_high_risk = len(user_metrics[user_metrics['Risk_Score'] >= 10])
            responsibles_with_critical = len(user_metrics[user_metrics['Critical_Actions'] > 0])
            total_value_managed = user_metrics['Total_Value'].sum()
            avg_actions_per_responsible = user_metrics['Total_Actions'].mean()
            avg_days_all_responsibles = user_metrics['Avg_Days'].mean()
            total_emitido = dashboard.df_enriched['Total_ValQtyIssued'].sum() if dashboard.df_enriched is not None else 0
            user_header = ft.Container(
                content=ft.Column([
                    ft.Text("👥 Análisis por Responsable", size=28, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text(f"Análisis detallado de {total_responsibles} responsables con acciones pendientes",
                            size=16, color=COLORS['text_secondary']),
                    ft.Row([
                        ft.ElevatedButton(
                            "Descargar a Excel",
                            icon=ft.Icons.DOWNLOAD,
                            bgcolor=COLORS['accent'],
                            color=COLORS['text_primary'],
                            on_click=descargar_excel
                        )
                    ], alignment=ft.MainAxisAlignment.END)
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )
            user_summary_metrics = ft.Row([
                create_compact_metric_card(
                    "Total Responsables",
                    f"{total_responsibles:,}",
                    "Con acciones asignadas",
                    COLORS['accent'],
                    "👥"
                ),
                create_compact_metric_card(
                    "Alto Riesgo",
                    f"{responsibles_high_risk:,}",
                    f"{(responsibles_high_risk/total_responsibles*100):.1f}% del total",
                    COLORS['error'],
                    "⚠️"
                ),
                create_compact_metric_card(
                    "Con Acciones Críticas",
                    f"{responsibles_with_critical:,}",
                    "Requieren atención inmediata",
                    COLORS['warning'],
                    "🔥"
                ),
                create_compact_metric_card(
                    "Promedio Acciones",
                    f"{avg_actions_per_responsible:.1f}",
                    f"Valor total: ${total_value_managed:,.0f}",
                    COLORS['success'],
                    "📊"
                ),
            ], wrap=True, spacing=15)
            additional_metrics = ft.Row([
                create_compact_metric_card(
                    "Valor Total Gestionado",
                    f"${total_value_managed:,.0f}",
                    "Por todos los responsables",
                    COLORS['success'],
                    "💰"
                ),
                create_compact_metric_card(
                    "Total Emitido",
                    f"${total_emitido:,.0f}",
                    "Suma de Total_ValQtyIssued",
                    COLORS['success'],
                    "💸"
                ),
                create_compact_metric_card(
                    "Promedio Días Pendientes",
                    f"{avg_days_all_responsibles:.0f}",
                    "Todos los responsables",
                    COLORS['warning'],
                    "⏰"
                ),
                
                
            ], wrap=True, spacing=15)
            for _, user in user_metrics.head(15).iterrows():
                if user['Ranking'] <= 3:
                    rank_color = COLORS['error']
                elif user['Ranking'] <= 10:
                    rank_color = COLORS['warning']
                else:
                    rank_color = COLORS['success']
                if user['Risk_Score'] >= 15:
                    risk_color = COLORS['error']
                elif user['Risk_Score'] >= 8:
                    risk_color = COLORS['warning']
                else:
                    risk_color = COLORS['success']
                user_metrics_table.rows.append(
                    ft.DataRow(
                        cells=[
                            ft.DataCell(ft.Text(f"#{user['Ranking']}", color=rank_color, weight=ft.FontWeight.BOLD, selectable=True)),
                            ft.DataCell(ft.Text(str(user['Responsible'])[:25] + "..." if len(str(user['Responsible'])) > 25 else str(user['Responsible']),
                                            color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(user['Email']), color=COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Total_Actions']:,}", color=COLORS['accent'], weight=ft.FontWeight.BOLD, selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Unique_WOs']:,}", color=COLORS['success'], selectable=True)),
                            ft.DataCell(ft.Text(f"${user['Total_Value']:,.0f}", color=COLORS['success'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Critical_Actions']:,}", color=COLORS['error'] if user['Critical_Actions'] > 0 else COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Avg_Days']:.0f}d", color=COLORS['warning'] if user['Avg_Days'] > 90 else COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Max_Days']:.0f}d", color=COLORS['warning'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Risk_Score']:.1f}", color=risk_color, weight=ft.FontWeight.BOLD, selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Critical_Rate']:.1f}%", color=COLORS['warning'], selectable=True)),
                        ]
                    )
                )
            table_container = ft.Container(
                content=ft.Column([
                    ft.Text("📊 Resumen Detallado (Top 15)",
                            size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Container(
                        content=ft.Column([user_metrics_table], scroll=ft.ScrollMode.ALWAYS),
                        border=ft.border.all(1, COLORS['secondary']),
                        border_radius=8,
                        padding=15,
                        height=400,
                        bgcolor=COLORS['primary'],
                    )
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )
            user_container.controls.extend([
                user_header,
                ft.Container(
                    content=ft.Column([
                        ft.Text("📈 Resumen Ejecutivo",
                                size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                        user_summary_metrics,
                        additional_metrics
                    ], spacing=15),
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12
                ),
                table_container
            ])
            page.update()
        update_user_metrics()
        return user_container


    def create_analytics_tab(page, dashboard):
        analytics_container = ft.Column(spacing=20)

        # Métricas y cards
        analytics_metrics_row = ft.Row([], spacing=15, wrap=True)
        transition_cards_row = ft.Row([], spacing=8, alignment=ft.MainAxisAlignment.START)
        info_cards_row = ft.Row([], spacing=15, wrap=True)

        # Tablas
        result_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("WONo", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("SO_FCST", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Sub", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("ItemNumber", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Description", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("PlanType", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("OpnQ", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("ParentWO", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("CreateDt", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Srt", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("UserId", color=COLORS['text_primary'])),
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=10,
        )
        expedite_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("id", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("AC", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("ShipDate", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("DemandSource", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("DemandType", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Ref", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Sub", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("EntityGroup", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("ItemNo", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("ReqQty", color=COLORS['text_primary'])),
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=10,
        )
        # Tabla de acciones históricas para el recuadro rojo
        actions_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("WO_Number", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Item_Number", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("PlanTp", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Act_Qty", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Std_Cost", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Total_ValQtyIssued", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Responsible_Final", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Días_Sin_Ejecutar", color=COLORS['text_primary'])),
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=10,
        )
        pr561_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("ItemNo", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("FuseNo", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("PlnType", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Srt", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("QtyOh", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("QtyIssue", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("QtyPending", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("ValQtyIss", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("ReqQty", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("WONo", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("ReqDate", color=COLORS['text_primary'])),
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=10,
        )

        # Contenedores de tablas
        tabla_container = ft.Container(
            content=ft.Column([
                ft.Text("📑 WO Liberadas", size=18, weight=ft.FontWeight.BOLD, color=COLORS['accent']),
                result_table
            ]),
            bgcolor=COLORS['surface'],
            padding=12,
            border_radius=12,
        )

        expedite_container = ft.Container(
            content=ft.Column([
                ft.Text("🚚 Expedites Relacionado con el Item", size=18, weight=ft.FontWeight.BOLD, color=COLORS['accent']),
                expedite_table
            ]),
            bgcolor=COLORS['surface'],
            padding=12,
            border_radius=12,
        )

        actions_container = ft.Container(
            content=ft.Column([
                ft.Text("🗂 Todas las Acciones Abiertas para el Item", size=18, weight=ft.FontWeight.BOLD, color=COLORS['accent']),
                actions_table
            ]),
            bgcolor=COLORS['surface'],
            padding=12,
            border_radius=12,
            height=220,
        )

        pr561_container = ft.Container(
            content=ft.Column([
                ft.Text("📦 Detalle de materiales", size=18, weight=ft.FontWeight.BOLD, color=COLORS['accent']),
                ft.Container(content=pr561_table)
            ]),
            bgcolor=COLORS['surface'],
            padding=12,
            border_radius=12,
        )

        # Campo de texto y botón
        search_wo = ft.TextField(
            label="Buscar WO",
            hint_text="Escribe o selecciona el número de WO...",
            autofocus=True,
            width=320,
            prefix_icon=ft.Icons.SEARCH,
        )

        def buscar_wo(e=None):
            print("➡️ Entrando a buscar_wo...") 
            result_table.rows.clear()
            expedite_table.rows.clear()
            actions_table.rows.clear()
            pr561_table.rows.clear()
            analytics_metrics_row.controls.clear()
            transition_cards_row.controls.clear()
            info_cards_row.controls.clear()
            # Si tienes mensaje de status, también límpialo:
            # status_message.value = ""
            page.update() 


            wo_number = search_wo.value.strip()
            if not wo_number:
                analytics_container.controls.append(ft.Text("Escribe un WO para buscar.", color=COLORS['error']))
                page.update()
                return

            # === 1. Consulta de WOInquiry ===
            df_result = get_wo_inquiry_parts(wo_number, dashboard.db_path)
            if len(df_result) == 0:
                analytics_container.controls.append(ft.Text("No se encontraron registros para ese WO.", color=COLORS['error']))
                page.update()
                return

            # --- MÉTRICAS CALCULADAS ---
            total_items = len(df_result)
            total_piezas = df_result['OpnQ'].astype(float).sum() if "OpnQ" in df_result.columns else 0

            # Toma el ItemNumber principal (de la primera WO)
            first_item_number = df_result.iloc[0]["ItemNumber"]

            # Total emitido de todas las WOs lanzadas con ese Item
            df_cancel = dashboard.df_original.copy()
            cancel_rows = df_cancel[df_cancel["Item_Number"] == first_item_number]
            total_emitido = cancel_rows['Total_ValQtyIssued'].replace(['$', ','], '', regex=True).astype(float).sum() if not cancel_rows.empty else 0

            analytics_metrics_row.controls.extend([
                create_compact_metric_card("# de WO Abiertas", f"{total_items}", "Total rows", COLORS["accent"], "🧩"),
                create_compact_metric_card("Total de Piezas", f"{total_piezas:,.0f}", "Suma OpnQ", COLORS["success"], "🔢"),
                create_compact_metric_card("Total Surtido", f"${total_emitido:,.2f}", "Material Surtido", COLORS["warning"], "💸"),
            ])

            # --- 2. Cards de transición y responsable
            if not cancel_rows.empty:
                row = cancel_rows.iloc[0]
                transition_cards_row.controls.extend([
                    create_compact_metric_card("Fm_Qty", f"{int(row['Fm_Qty'])}", "Cantidad Inicial", COLORS["success"], "🏁"),
                    ft.Text("➡️", size=28, color=COLORS["accent"], weight=ft.FontWeight.BOLD),
                    create_compact_metric_card("To_Qty", f"{int(row['To_Qty'])}", "Cantidad Final", COLORS["warning"], "🏁"),
                ])
                info_cards_row.controls.extend([
                    create_compact_metric_card("Días Sin Ejecutar", f"{int(row['Días_Sin_Ejecutar'])}", "", COLORS["error"], "⏰"),
                    create_compact_metric_card("Responsable", f"{row['Responsible_Final']}", "", COLORS["accent"], "👤"),
                ])
            else:
                transition_cards_row.controls.clear()
                info_cards_row.controls.clear()

            # --- 3. Llenar la tabla de partes WO (todas las celdas copiables) ---
            for _, row in df_result.iterrows():
                result_table.rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(row["WONo"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["SO_FCST"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Sub"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["ItemNumber"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Description"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["PlanType"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["OpnQ"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["ParentWO"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["CreateDt"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Srt"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["UserId"]), selectable=True)),
                ]))

            # --- 4. Expedites ---
            conn = sqlite3.connect(dashboard.db_path)
            df_expedite = pd.read_sql_query("""
                SELECT id, AC, ShipDate, DemandSource, DemandType, Ref, Sub, EntityGroup, ItemNo, ReqQty
                FROM expedite
                WHERE ItemNo = ?
            """, conn, params=(first_item_number,))
            conn.close()
            total_req_qty = df_expedite['ReqQty'].astype(float).sum() if not df_expedite.empty else 0
            analytics_metrics_row.controls.append(
                create_compact_metric_card(
                    "Total de Requerimiento",
                    f"{total_req_qty:,.0f}",
                    "Suma ReqQty (Expedite)",
                    COLORS["warning"],
                    "📦"
                )
            )
            for _, row in df_expedite.iterrows():
                expedite_table.rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(row["id"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["AC"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["ShipDate"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["DemandSource"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["DemandType"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Ref"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Sub"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["EntityGroup"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["ItemNo"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["ReqQty"]), selectable=True)),
                ]))

            # --- 5. Acciones (todas las cancelaciones para ese Item) ---
            for _, row in cancel_rows.iterrows():
                actions_table.rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(row["WO_Number"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Item_Number"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["PlanTp"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Act_Qty"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Std_Cost"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Total_ValQtyIssued"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Responsible_Final"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Días_Sin_Ejecutar"]), selectable=True)),
                ]))
            # --- 6. Materiales de la WO seleccionada (PR561) ---
            # selected_wo = df_result.iloc[0]["WONo"]
            conn = sqlite3.connect(dashboard.db_path)
            df_pr561 = pd.read_sql_query("""
                SELECT id, Entity, ItemNo, FuseNo, PlnType, Srt, St, QtyOh, QtyIssue, QtyPending, ReqQty, ValQtyIss, ValNotIss, ValRequired, WONo, WODescripton, ReqDate
                FROM pr561
                WHERE WONo = ?
            """, conn, params=(wo_number,))
            conn.close()
            print(df_pr561)
            print("Total rows encontrados:", len(df_pr561))
            if df_pr561.empty:
                pr561_warning = ft.Container(
                    content=ft.Text(
                        "⚠️ Validar: WO cerrada, no tiene materiales asignados.",
                        size=18,
                        color=COLORS['warning'],
                        weight=ft.FontWeight.BOLD,
                    ),
                    padding=ft.padding.all(16),
                    bgcolor=COLORS['surface'],
                    border_radius=10,
                )
                pr561_container.content.controls.append(pr561_warning)
            else:
                for _, row in df_pr561.iterrows():
                    pr561_table.rows.append(
                        ft.DataRow(cells=[
                            # ... tus columnas ...
                        ])
                    )
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            for _, row in df_pr561.iterrows():
                pr561_table.rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(row["ItemNo"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["FuseNo"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["PlnType"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["Srt"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["QtyOh"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["QtyIssue"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["QtyPending"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["ValQtyIss"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["ReqQty"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["WONo"]), selectable=True)),
                    ft.DataCell(ft.Text(str(row["ReqDate"]), selectable=True)),
                ]))
            page.update()

        search_button = ft.ElevatedButton(
            "Buscar WO",
            icon=ft.Icons.SEARCH,
            bgcolor=COLORS['accent'],
            color=COLORS['text_primary'],
            on_click=buscar_wo,
        )

        analytics_container.controls = [
            ft.Container(
                content=ft.Text(
                    "🔬 Análisis Avanzado de WO a Cancelar",
                    size=26,
                    weight=ft.FontWeight.BOLD,
                    color=COLORS['accent'],
                ),
                padding=ft.padding.only(bottom=10),
            ),
            ft.Row([search_wo, search_button], spacing=10),
            analytics_metrics_row,
            ft.Row([transition_cards_row, info_cards_row], spacing=25),
            ft.Divider(),
            tabla_container,
            expedite_container,
            actions_container,
            pr561_container
        ]
        return analytics_container


    # Tabs solo: Por Responsable y Analisis Avanzado
    tabs = ft.Tabs(
        selected_index=0,
        animation_duration=300,
        tabs=[
            ft.Tab(
                text="👥 Por Responsable",
                content=create_user_metrics_tab(),
                icon=ft.Icons.PEOPLE
            ),
            ft.Tab(
                text="🔬 Análisis Avanzado",
                content=create_analytics_tab(page, dashboard),
                icon=ft.Icons.ANALYTICS
            )
        ],
        expand=1,
        label_color=COLORS['text_secondary'],
        indicator_color=COLORS['accent']
    )
    page.add(
        ft.Container(
            content=ft.Text(
                "WO Cancellation Actions Dashboard",
                size=34,
                weight=ft.FontWeight.BOLD,
                color=COLORS['accent'],
            ),
            padding=ft.padding.symmetric(vertical=12),
            alignment=ft.alignment.center,
        )
    )
    page.add(tabs)

def get_wo_inquiry_parts(wo_number, db_path):
    conn = sqlite3.connect(db_path)
    # 1. Buscar el ItemNumber de la WO consultada
    df_wo = pd.read_sql_query("""
        SELECT ItemNumber FROM WOInquiry WHERE WONo = ?
    """, conn, params=(wo_number,))
    if df_wo.empty:
        conn.close()
        return pd.DataFrame()  # No existe esa WO
    
    item_number = df_wo.iloc[0]['ItemNumber']
    # 2. Traer TODAS las WOs que tengan ese mismo ItemNumber
    df = pd.read_sql_query("""
        SELECT * FROM WOInquiry WHERE ItemNumber = ?
    """, conn, params=(item_number,))
    conn.close()
    return df

def get_expedite_parts(item_numbers, db_path):
    conn = sqlite3.connect(db_path)
    # Asegúrate que item_numbers es una lista/serie de strings
    placeholders = ','.join(['?'] * len(item_numbers))
    query = f"""
        SELECT id, AC, ShipDate, DemandSource, DemandType, Ref, Sub, EntityGroup, ItemNo, ReqQty
        FROM expedite
        WHERE ItemNo IN ({placeholders})
    """
    df = pd.read_sql_query(query, conn, params=list(item_numbers))
    conn.close()
    return df

if __name__ == "__main__":
    ft.app(target=main)