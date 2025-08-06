import flet as ft
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import numpy as np

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
        # Ruta de base de datos
        self.db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        self.df_original = None
        self.df_enriched = None
        
    def load_and_process_data(self):
        """Cargar y procesar datos de la vista Cancelation_Information_WO"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            # Query de la vista con las columnas reales
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
            
            # Procesar datos
            self.process_enriched_data()
            
            print(f"✅ Datos procesados: {len(self.df_original)} WOs con acciones de cancelación")
            return True
            
        except Exception as e:
            print(f"❌ Error cargando datos de cancelaciones: {e}")
            return False
    
    def process_enriched_data(self):
        """Procesar datos enriquecidos de WOs con acciones de cancelación"""
        df = self.df_original.copy()
        
        # Convertir fecha de primera aparición
        df['Fecha_Primera_Aparicion'] = pd.to_datetime(df['Fecha_Primera_Aparicion'], errors='coerce')
        
        # Convertir columnas numéricas
        numeric_columns = ['Días_Sin_Ejecutar', 'Std_Cost', 'Act_Qty', 'Fm_Qty', 'To_Qty', 'Total_ValQtyIssued']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Calcular métricas derivadas
        current_date = datetime.now()
        
        # Valor total de la acción (Act_Qty * Std_Cost)
        df['Action_Value'] = df['Act_Qty'] * df['Std_Cost']
        
        # Valor total potencial (incluyendo materiales emitidos)
        df['Total_Potential_Value'] = df['Action_Value'] + df['Total_ValQtyIssued']
        
        # Categorización por días sin ejecutar
        df['Aging_Category'] = df['Días_Sin_Ejecutar'].apply(self._categorize_aging)
        
        # Categorización por valor de la acción
        df['Value_Category'] = df['Action_Value'].apply(self._categorize_value)
        
        # Categorización por tipo de plan
        df['Plan_Category'] = df['PlanTp'].apply(self._categorize_plan_type)
        
        # Estado del responsable
        df['Responsible_Status'] = df.apply(self._determine_responsible_status, axis=1)
        
        # Nivel de criticidad basado en días sin ejecutar y valor
        df['Criticality_Level'] = df.apply(self._determine_criticality, axis=1)
        
        # Limpiar datos de responsables
        df['Responsible_Final'] = df['Responsible_Final'].fillna('Sin Asignar')
        df['Email_Final'] = df['Email_Final'].fillna('No Email')
        
        self.df_enriched = df
    
    def _categorize_aging(self, days):
        """Categorizar por días sin ejecutar"""
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
        """Categorizar por valor de la acción"""
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
        """Categorizar tipos de plan"""
        if pd.isna(plan_type) or plan_type == '':
            return 'Sin Tipo'
        
        # Mapear tipos comunes de plan
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
        """Determinar estado del responsable"""
        if pd.notna(row['Responsible_Final']) and row['Responsible_Final'] != 'Sin Asignar':
            if pd.notna(row['Email_Final']) and row['Email_Final'] != 'No Email':
                return 'Completo'
            else:
                return 'Sin Email'
        else:
            return 'Sin Asignar'
    
    def _determine_criticality(self, row):
        """Determinar nivel de criticidad"""
        days = row['Días_Sin_Ejecutar']
        value = row['Action_Value']
        
        if pd.isna(days):
            days = 0
        if pd.isna(value):
            value = 0
        
        # Scoring de criticidad
        score = 0
        
        # Puntos por días sin ejecutar
        if days > 180:
            score += 4
        elif days > 90:
            score += 3
        elif days > 30:
            score += 2
        elif days > 7:
            score += 1
        
        # Puntos por valor
        if value > 10000:
            score += 3
        elif value > 1000:
            score += 2
        elif value > 100:
            score += 1
        
        # Clasificación final
        if score >= 6:
            return 'Crítico'
        elif score >= 4:
            return 'Alto'
        elif score >= 2:
            return 'Medio'
        else:
            return 'Bajo'
    
    def get_general_metrics(self):
        """Calcular métricas generales"""
        if self.df_enriched is None or len(self.df_enriched) == 0:
            return self._get_empty_general_metrics()
        
        df = self.df_enriched
        
        # Métricas básicas
        total_wo_actions = len(df)
        unique_wos = df['WO_Number'].nunique()
        unique_entities = df['Entity_WorkOrder'].nunique()
        
        # Métricas temporales (aging)
        aging_counts = df['Aging_Category'].value_counts()
        critical_aging = aging_counts.get('Crítico (>6M)', 0)
        six_months = aging_counts.get('6 Meses', 0)
        three_months = aging_counts.get('3 Meses', 0)
        one_month = aging_counts.get('1 Mes', 0)
        
        # Métricas de valor
        total_action_value = df['Action_Value'].sum()
        total_potential_value = df['Total_Potential_Value'].sum()
        avg_action_value = df['Action_Value'].mean()
        
        # Métricas de criticidad
        criticality_counts = df['Criticality_Level'].value_counts()
        critical_level = criticality_counts.get('Crítico', 0)
        high_level = criticality_counts.get('Alto', 0)
        
        # Métricas de responsables
        with_responsible = len(df[df['Responsible_Status'] == 'Completo'])
        without_responsible = len(df[df['Responsible_Status'] == 'Sin Asignar'])
        unique_responsibles = len(df[df['Responsible_Final'] != 'Sin Asignar']['Responsible_Final'].unique())
        
        # Promedio de días sin ejecutar
        avg_days_pending = df['Días_Sin_Ejecutar'].mean()
        max_days_pending = df['Días_Sin_Ejecutar'].max()
        
        # Top entity y plan type
        top_entity = df['Entity_WorkOrder'].mode().iloc[0] if len(df['Entity_WorkOrder'].mode()) > 0 else 'N/A'
        top_plan_type = df['Plan_Category'].mode().iloc[0] if len(df['Plan_Category'].mode()) > 0 else 'N/A'
        
        return {
            'total_wo_actions': total_wo_actions,
            'unique_wos': unique_wos,
            'unique_entities': unique_entities,
            'critical_aging': critical_aging,
            'six_months': six_months,
            'three_months': three_months,
            'one_month': one_month,
            'total_action_value': total_action_value,
            'total_potential_value': total_potential_value,
            'avg_action_value': avg_action_value,
            'critical_level': critical_level,
            'high_level': high_level,
            'with_responsible': with_responsible,
            'without_responsible': without_responsible,
            'unique_responsibles': unique_responsibles,
            'avg_days_pending': avg_days_pending,
            'max_days_pending': max_days_pending,
            'top_entity': top_entity,
            'top_plan_type': top_plan_type
        }
    
    def _get_empty_general_metrics(self):
        """Retornar métricas vacías"""
        return {
            'total_wo_actions': 0,
            'unique_wos': 0,
            'unique_entities': 0,
            'critical_aging': 0,
            'six_months': 0,
            'three_months': 0,
            'one_month': 0,
            'total_action_value': 0,
            'total_potential_value': 0,
            'avg_action_value': 0,
            'critical_level': 0,
            'high_level': 0,
            'with_responsible': 0,
            'without_responsible': 0,
            'unique_responsibles': 0,
            'avg_days_pending': 0,
            'max_days_pending': 0,
            'top_entity': 'N/A',
            'top_plan_type': 'N/A'
        }
    
    def get_user_metrics(self):
        """Generar métricas detalladas por responsable"""
        if self.df_enriched is None or len(self.df_enriched) == 0:
            return pd.DataFrame()
        
        df = self.df_enriched
        
        # Filtrar solo registros con responsable asignado
        df_with_users = df[
            (df['Responsible_Final'].notna()) & 
            (df['Responsible_Final'] != 'Sin Asignar') & 
            (df['Responsible_Final'] != '')
        ].copy()
        
        if len(df_with_users) == 0:
            return pd.DataFrame()
        
        # Agrupar por responsable y calcular KPIs
        user_metrics = df_with_users.groupby(['Responsible_Final', 'Email_Final']).agg({
            'WO_Number': ['count', 'nunique'],  # Total acciones y WOs únicos
            'Action_Value': ['sum', 'mean'],  # Valor total y promedio
            'Total_Potential_Value': 'sum',  # Valor potencial total
            'Días_Sin_Ejecutar': ['mean', 'max'],  # Promedio y máximo días
            'Criticality_Level': [
                lambda x: (x == 'Crítico').sum(),
                lambda x: (x == 'Alto').sum(),
                lambda x: (x == 'Medio').sum()
            ],
            'Aging_Category': lambda x: (x == 'Crítico (>6M)').sum(),
            'Entity_WorkOrder': 'nunique'  # Entidades únicas
        }).reset_index()
        
        # Aplanar columnas multinivel
        user_metrics.columns = [
            'Responsible', 'Email', 'Total_Actions', 'Unique_WOs', 'Total_Value', 'Avg_Value',
            'Total_Potential', 'Avg_Days', 'Max_Days', 'Critical_Actions',
            'High_Actions', 'Medium_Actions', 'Critical_Aging', 'Unique_Entities'
        ]
        
        # Calcular métricas adicionales
        user_metrics['Risk_Score'] = (
            (user_metrics['Critical_Actions'] * 5) + 
            (user_metrics['High_Actions'] * 3) + 
            (user_metrics['Critical_Aging'] * 4) +
            (user_metrics['Avg_Days'] / 30)  # Factor de días promedio
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
        
        # Ordenar por risk score descendente
        user_metrics = user_metrics.sort_values('Risk_Score', ascending=False).reset_index(drop=True)
        user_metrics['Ranking'] = range(1, len(user_metrics) + 1)
        
        return user_metrics

def create_compact_metric_card(title, value, subtitle="", color=COLORS['accent'], icon="📊"):
    """Crear tarjeta de métrica compacta y ajustada al contenido"""
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
        width=180,  # Ancho fijo compacto
        height=120   # Altura fija compacta
    )

def main(page: ft.Page):
    page.title = "WO Cancellation Actions Dashboard"
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = COLORS['surface']
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20
    
    dashboard = WOCancellationDashboard()
    
    def create_general_tab():
        """Tab de métricas generales de acciones de cancelación pendientes"""
        main_container = ft.Column(spacing=20)
        status_text = ft.Text("Iniciando análisis de acciones de cancelación...", color=COLORS['text_secondary'])
        
        # Tabla de acciones pendientes
        actions_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("WO Number", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Días Pendiente", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Primera Aparición", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Entity", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Item Description", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Plan Type", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Valor Acción", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Valor Potencial", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Responsible", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Criticality", color=COLORS['text_primary']))
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=8,
        )
        
        def update_general_dashboard():
            main_container.controls.clear()
            actions_table.rows.clear()
            
            status_text.value = "🔄 Cargando datos de acciones de cancelación..."
            status_text.color = COLORS['accent']
            page.update()
            
            if dashboard.load_and_process_data():
                metrics = dashboard.get_general_metrics()
                
                if metrics['total_wo_actions'] == 0:
                    main_container.controls.append(
                        ft.Container(
                            content=ft.Column([
                                ft.Text("ℹ️ No hay acciones de cancelación pendientes", 
                                       size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                ft.Text("No se encontraron registros en Cancelation_Information_WO", 
                                       size=16, color=COLORS['text_secondary']),
                                ft.ElevatedButton(
                                    "🔄 Actualizar Datos",
                                    on_click=lambda _: update_general_dashboard(),
                                    bgcolor=COLORS['accent'],
                                    color=COLORS['text_primary'],
                                    icon=ft.Icons.REFRESH
                                ),
                            ], spacing=20, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                            padding=ft.padding.all(40),
                            bgcolor=COLORS['primary'],
                            border_radius=12
                        )
                    )
                    status_text.value = "⚠️ No hay datos para mostrar"
                    status_text.color = COLORS['warning']
                    page.update()
                    return
                
                # Header ejecutivo
                header = ft.Container(
                    content=ft.Row([
                        ft.Column([
                            ft.Text("⏰ WO Cancellation Actions Dashboard", 
                                size=32, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            ft.Text(f"Análisis de {metrics['total_wo_actions']:,} acciones de cancelación pendientes", 
                                size=16, color=COLORS['text_secondary']),
                            ft.Text("Fuente: Cancelation_Information_WO (cancel_action_aging)", 
                                size=12, color=COLORS['warning'])
                        ], expand=True),
                        ft.Column([
                            ft.Row([
                                ft.ElevatedButton(
                                    "🔄 Actualizar Datos",
                                    on_click=lambda _: update_general_dashboard(),
                                    bgcolor=COLORS['accent'],
                                    color=COLORS['text_primary'],
                                    icon=ft.Icons.REFRESH
                                ),
                            ], spacing=10),
                            status_text
                        ], horizontal_alignment=ft.CrossAxisAlignment.END)
                    ]),
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12
                )
                
                # Métricas principales en 4 columnas compactas
                main_metrics_grid = ft.Row([
                    create_compact_metric_card(
                        "Acciones Pendientes",
                        f"{metrics['total_wo_actions']:,}",
                        f"En {metrics['unique_wos']:,} WOs",
                        COLORS['error'],
                        "⏰"
                    ),
                    create_compact_metric_card(
                        "Crítico (>6M)",
                        f"{metrics['critical_aging']:,}",
                        f"Max: {metrics['max_days_pending']:.0f}d",
                        COLORS['error'],
                        "🚨"
                    ),
                    create_compact_metric_card(
                        "Nivel Crítico",
                        f"{metrics['critical_level']:,}",
                        f"Alto: {metrics['high_level']:,}",
                        COLORS['warning'],
                        "⚠️"
                    ),
                    create_compact_metric_card(
                        "Promedio Días",
                        f"{metrics['avg_days_pending']:.0f}",
                        f"{metrics['unique_entities']:,} entidades",
                        COLORS['accent'],
                        "📊"
                    ),
                ], spacing=15, wrap=True, alignment=ft.MainAxisAlignment.START)
                
                # Métricas de valor en 4 columnas compactas
                value_metrics_grid = ft.Row([
                    create_compact_metric_card(
                        "Valor Total",
                        f"${metrics['total_action_value']:,.0f}",
                        f"Prom: ${metrics['avg_action_value']:,.0f}",
                        COLORS['success'],
                        "💰"
                    ),
                    create_compact_metric_card(
                        "Valor Potencial",
                        f"${metrics['total_potential_value']:,.0f}",
                        "Inc. materiales",
                        COLORS['success'],
                        "💎"
                    ),
                    create_compact_metric_card(
                        "Con Responsable",
                        f"{metrics['with_responsible']:,}",
                        f"Sin: {metrics['without_responsible']:,}",
                        COLORS['success'] if metrics['with_responsible'] > metrics['without_responsible'] else COLORS['error'],
                        "👥"
                    ),
                    create_compact_metric_card(
                        "Responsables",
                        f"{metrics['unique_responsibles']:,}",
                        f"{metrics['top_entity'][:8]}..." if len(metrics['top_entity']) > 8 else metrics['top_entity'],
                        COLORS['accent'],
                        "🎯"
                    ),
                ], spacing=15, wrap=True, alignment=ft.MainAxisAlignment.START)
                
                # Métricas de aging en 4 columnas compactas
                aging_metrics_grid = ft.Row([
                    create_compact_metric_card(
                        "1 Mes",
                        f"{metrics['one_month']:,}",
                        "Pendientes",
                        COLORS['warning'],
                        "📅"
                    ),
                    create_compact_metric_card(
                        "3 Meses",
                        f"{metrics['three_months']:,}",
                        "Atención",
                        COLORS['error'],
                        "📆"
                    ),
                    create_compact_metric_card(
                        "6 Meses",
                        f"{metrics['six_months']:,}",
                        "Alto riesgo",
                        COLORS['error'],
                        "🗓️"
                    ),
                ], spacing=15, wrap=True, alignment=ft.MainAxisAlignment.START)
                
                # Llenar tabla con acciones pendientes
                if dashboard.df_enriched is not None:
                    for _, row in dashboard.df_enriched.head(50).iterrows():
                        # Determinar colores según criticidad
                        criticality = row.get('Criticality_Level', 'Bajo')
                        if criticality == 'Crítico':
                            crit_color = COLORS['error']
                        elif criticality == 'Alto':
                            crit_color = COLORS['warning']
                        elif criticality == 'Medio':
                            crit_color = COLORS['accent']
                        else:
                            crit_color = COLORS['success']
                        
                        # Color por días pendientes
                        days = row.get('Días_Sin_Ejecutar', 0)
                        if days > 180:
                            days_color = COLORS['error']
                        elif days > 90:
                            days_color = COLORS['warning']
                        else:
                            days_color = COLORS['text_primary']
                        
                        # Color del valor
                        value = row.get('Action_Value', 0)
                        if value > 10000:
                            value_color = COLORS['error']
                        elif value > 1000:
                            value_color = COLORS['warning']
                        else:
                            value_color = COLORS['text_primary']
                        
                        actions_table.rows.append(
                            ft.DataRow(
                                cells=[
                                    ft.DataCell(ft.Text(str(row.get('WO_Number', '')), color=COLORS['accent'], selectable=True)),
                                    ft.DataCell(ft.Text(f"{row.get('Días_Sin_Ejecutar', 0):.0f}", color=days_color, weight=ft.FontWeight.BOLD, selectable=True)),
                                    ft.DataCell(ft.Text(str(row.get('Fecha_Primera_Aparicion', ''))[:10] if pd.notna(row.get('Fecha_Primera_Aparicion')) else "-", 
                                              color=COLORS['text_secondary'], selectable=True)),
                                    ft.DataCell(ft.Text(str(row.get('Entity_WorkOrder', ''))[:8] + "..." if len(str(row.get('Entity_WorkOrder', ''))) > 8 else str(row.get('Entity_WorkOrder', '')), 
                                              color=COLORS['text_secondary'], selectable=True)),
                                    ft.DataCell(ft.Text(str(row.get('Item_Description', ''))[:25] + "..." if len(str(row.get('Item_Description', ''))) > 25 else str(row.get('Item_Description', '')), 
                                              color=COLORS['text_primary'], selectable=True)),
                                    ft.DataCell(ft.Text(str(row.get('PlanTp', '')), color=COLORS['text_secondary'], selectable=True)),
                                    ft.DataCell(ft.Text(f"${row.get('Action_Value', 0):,.0f}", color=value_color, selectable=True)),
                                    ft.DataCell(ft.Text(f"${row.get('Total_Potential_Value', 0):,.0f}", color=COLORS['success'], selectable=True)),
                                    ft.DataCell(ft.Text(str(row.get('Responsible_Final', ''))[:15] + "..." if len(str(row.get('Responsible_Final', ''))) > 15 else str(row.get('Responsible_Final', '')), 
                                              color=COLORS['text_secondary'], selectable=True)),
                                    ft.DataCell(ft.Text(criticality, color=crit_color, weight=ft.FontWeight.BOLD, selectable=True)),
                                ]
                            )
                        )
                
                # Contenedor de tabla
                table_container = ft.Container(
                    content=ft.Column([
                        ft.Text("📋 Acciones de Cancelación Pendientes (Top 50)", 
                               size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                        ft.Text("Ordenadas por días sin ejecutar - Celdas seleccionables", 
                               size=12, color=COLORS['text_secondary']),
                        ft.Container(
                            content=ft.Column([actions_table], scroll=ft.ScrollMode.ALWAYS),
                            border=ft.border.all(1, COLORS['secondary']),
                            border_radius=8,
                            padding=15,
                            height=500,
                            bgcolor=COLORS['primary'],
                        )
                    ]),
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12
                )
                
                # Agregar todos los componentes con layout mejorado
                main_container.controls.extend([
                    header,
                    ft.Container(
                        content=ft.Column([
                            ft.Text("📊 Métricas Principales", 
                                   size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            main_metrics_grid,
                        ], spacing=15),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=12
                    ),
                    ft.Container(
                        content=ft.Column([
                            ft.Text("💰 Valor e Impacto", 
                                   size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            value_metrics_grid,
                        ], spacing=15),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=12
                    ),
                    ft.Container(
                        content=ft.Column([
                            ft.Text("⏳ Análisis de Aging", 
                                   size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            aging_metrics_grid,
                        ], spacing=15),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=12
                    ),
                    table_container
                ])
                
                status_text.value = "✅ Dashboard cargado exitosamente"
                status_text.color = COLORS['success']
            else:
                status_text.value = "❌ Error cargando datos"
                status_text.color = COLORS['error']
            
            page.update()
        
        update_general_dashboard()
        return main_container
    
    def create_user_metrics_tab():
        """Tab de métricas por responsable"""
        user_container = ft.Column(spacing=20)

        # Tabla de métricas por responsable
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

        def update_user_metrics():
            user_container.controls.clear()
            user_metrics_table.rows.clear()

            if dashboard.df_enriched is None:
                user_container.controls.append(
                    ft.Text("⚠️ Carga primero el dashboard principal", size=16, color=COLORS['warning'])
                )
                page.update()
                return

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

            # Métricas resumen de responsables
            total_responsibles = len(user_metrics)
            responsibles_high_risk = len(user_metrics[user_metrics['Risk_Score'] >= 10])
            responsibles_with_critical = len(user_metrics[user_metrics['Critical_Actions'] > 0])
            total_value_managed = user_metrics['Total_Value'].sum()
            avg_actions_per_responsible = user_metrics['Total_Actions'].mean()
            avg_days_all_responsibles = user_metrics['Avg_Days'].mean()

            # Header para métricas de responsables
            user_header = ft.Container(
                content=ft.Column([
                    ft.Text("👥 Análisis por Responsable", size=28, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text(f"Análisis detallado de {total_responsibles} responsables con acciones pendientes",
                            size=16, color=COLORS['text_secondary']),
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )

            # Métricas de responsables (usando create_compact_metric_card)
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

            # Métricas adicionales de responsables
            additional_metrics = ft.Row([
                create_compact_metric_card(
                    "Valor Total Gestionado",
                    f"${total_value_managed:,.0f}",
                    "Por todos los responsables",
                    COLORS['success'],
                    "💰"
                ),
                create_compact_metric_card(
                    "Promedio Días Pendientes",
                    f"{avg_days_all_responsibles:.0f}",
                    "Todos los responsables",
                    COLORS['warning'],
                    "⏰"
                ),
                create_compact_metric_card(
                    "Cobertura Responsables",
                    f"{(dashboard.get_general_metrics()['with_responsible']/dashboard.get_general_metrics()['total_wo_actions']*100):.1f}%",
                    "Acciones con responsable",
                    COLORS['accent'],
                    "📈"
                ),
                create_compact_metric_card(
                    "Eficiencia Promedio",
                    f"{user_metrics['Efficiency_Score'].mean():.1f}%",
                    "Score de eficiencia",
                    COLORS['text_primary'],
                    "🎯"
                ),
            ], wrap=True, spacing=15)

            # Tabla compacta con datos detallados (Top 15)
            for _, user in user_metrics.head(15).iterrows():
                # Determinar colores según ranking y riesgo
                if user['Ranking'] <= 3:
                    rank_color = COLORS['error']
                elif user['Ranking'] <= 10:
                    rank_color = COLORS['warning']
                else:
                    rank_color = COLORS['success']

                # Color del risk score
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

            # Contenedor de tabla compacta
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








    def create_analytics_tab():
        """Tab para análisis avanzados de acciones de cancelación"""
        analytics_container = ft.Column(spacing=20)
        
        # Header del tab de analytics
        analytics_header = ft.Container(
            content=ft.Column([
                ft.Text("🔬 Análisis Avanzado de Cancelaciones", 
                       size=28, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Text("Análisis profundo de patrones y tendencias en acciones de cancelación", 
                       size=16, color=COLORS['text_secondary']),
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
        
        # Sección de análisis por entidad
        entity_analysis = ft.Container(
            content=ft.Column([
                ft.Text("🏢 Análisis por Entidad", 
                       size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Text("Distribución de acciones pendientes por entidad organizacional", 
                       size=14, color=COLORS['text_secondary']),
                ft.Container(
                    content=ft.Column([
                        ft.Card(
                            content=ft.Container(
                                padding=ft.padding.all(15),
                                content=ft.Column([
                                    ft.Text("📊 Métricas por Entidad", 
                                           size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("• Identificar entidades con mayor backlog de cancelaciones", 
                                           size=12, color=COLORS['text_secondary']),
                                    ft.Text("• Análisis de valor promedio por entidad", 
                                           size=12, color=COLORS['text_secondary']),
                                    ft.Text("• Tiempo promedio de resolución por entidad", 
                                           size=12, color=COLORS['text_secondary']),
                                ]),
                                bgcolor=COLORS['secondary']
                            )
                        ),
                    ]),
                    padding=ft.padding.all(10)
                )
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
        
        # Sección de análisis temporal
        temporal_analysis = ft.Container(
            content=ft.Column([
                ft.Text("📅 Análisis Temporal", 
                       size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Text("Tendencias y patrones temporales en acciones de cancelación", 
                       size=14, color=COLORS['text_secondary']),
                ft.Container(
                    content=ft.Column([
                        ft.Card(
                            content=ft.Container(
                                padding=ft.padding.all(15),
                                content=ft.Column([
                                    ft.Text("📈 Análisis de Tendencias", 
                                           size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("• Gráfico de aging por períodos", 
                                           size=12, color=COLORS['text_secondary']),
                                    ft.Text("• Identificación de picos estacionales", 
                                           size=12, color=COLORS['text_secondary']),
                                    ft.Text("• Predicción de acumulación futura", 
                                           size=12, color=COLORS['text_secondary']),
                                ]),
                                bgcolor=COLORS['secondary']
                            )
                        ),
                    ]),
                    padding=ft.padding.all(10)
                )
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
        
        # Sección de análisis de valor
        value_analysis = ft.Container(
            content=ft.Column([
                ft.Text("💰 Análisis de Impacto Financiero", 
                       size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Text("Impacto económico de acciones de cancelación pendientes", 
                       size=14, color=COLORS['text_secondary']),
                ft.Container(
                    content=ft.Row([
                        ft.Card(
                            content=ft.Container(
                                padding=ft.padding.all(15),
                                content=ft.Column([
                                    ft.Text("📊 Distribución de Valor", 
                                           size=14, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("Análisis Pareto de acciones por valor", 
                                           size=11, color=COLORS['text_secondary']),
                                ]),
                                bgcolor=COLORS['secondary']
                            )
                        ),
                        ft.Card(
                            content=ft.Container(
                                padding=ft.padding.all(15),
                                content=ft.Column([
                                    ft.Text("💸 Costo de Retraso", 
                                           size=14, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("Impacto financiero del aging", 
                                           size=11, color=COLORS['text_secondary']),
                                ]),
                                bgcolor=COLORS['secondary']
                            )
                        ),
                    ], spacing=10),
                    padding=ft.padding.all(10)
                )
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
        
        # Funcionalidades futuras
        future_features = ft.Container(
            content=ft.Column([
                ft.Text("🚀 Funcionalidades Futuras", 
                       size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Column([
                    ft.Card(
                        content=ft.Container(
                            padding=ft.padding.all(15),
                            content=ft.Row([
                                ft.Icon(ft.Icons.ALARM, color=COLORS['error'], size=30),
                                ft.Column([
                                    ft.Text("Sistema de Alertas", 
                                           size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("Notificaciones automáticas para acciones críticas", 
                                           size=12, color=COLORS['text_secondary'])
                                ], expand=True)
                            ]),
                            bgcolor=COLORS['secondary']
                        )
                    ),
                    ft.Card(
                        content=ft.Container(
                            padding=ft.padding.all(15),
                            content=ft.Row([
                                ft.Icon(ft.Icons.AUTO_GRAPH, color=COLORS['success'], size=30),
                                ft.Column([
                                    ft.Text("Machine Learning", 
                                           size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("Predicción de WOs con alta probabilidad de cancelación", 
                                           size=12, color=COLORS['text_secondary'])
                                ], expand=True)
                            ]),
                            bgcolor=COLORS['secondary']
                        )
                    ),
                    ft.Card(
                        content=ft.Container(
                            padding=ft.padding.all(15),
                            content=ft.Row([
                                ft.Icon(ft.Icons.DESCRIPTION, color=COLORS['accent'], size=30),
                                ft.Column([
                                    ft.Text("Reportes Automatizados", 
                                           size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("Generación automática de reportes ejecutivos", 
                                           size=12, color=COLORS['text_secondary'])
                                ], expand=True)
                            ]),
                            bgcolor=COLORS['secondary']
                        )
                    ),
                    ft.Card(
                        content=ft.Container(
                            padding=ft.padding.all(15),
                            content=ft.Row([
                                ft.Icon(ft.Icons.INTEGRATION_INSTRUCTIONS, color=COLORS['warning'], size=30),
                                ft.Column([
                                    ft.Text("Integración API", 
                                           size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("Conexión directa con sistemas de Work Orders", 
                                           size=12, color=COLORS['text_secondary'])
                                ], expand=True)
                            ]),
                            bgcolor=COLORS['secondary']
                        )
                    ),
                ], spacing=10)
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
        
        # Botón de configuración (placeholder)
        config_section = ft.Container(
            content=ft.Column([
                ft.Text("⚙️ Configuración Avanzada", 
                       size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Column([
                    ft.ElevatedButton(
                        "🔧 Configurar Umbrales de Criticidad",
                        bgcolor=COLORS['accent'],
                        color=COLORS['text_primary'],
                        disabled=True,
                        icon=ft.Icons.SETTINGS
                    ),
                    ft.ElevatedButton(
                        "📧 Configurar Notificaciones",
                        bgcolor=COLORS['warning'],
                        color=COLORS['text_primary'],
                        disabled=True,
                        icon=ft.Icons.NOTIFICATIONS
                    ),
                    ft.ElevatedButton(
                        "📊 Exportar Datos",
                        bgcolor=COLORS['success'],
                        color=COLORS['text_primary'],
                        disabled=True,
                        icon=ft.Icons.DOWNLOAD
                    ),
                ], spacing=10)
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
        
        analytics_container.controls.extend([
            analytics_header,
            entity_analysis,
            temporal_analysis,
            value_analysis,
            future_features,
            config_section
        ])
        
        return analytics_container

    # Crear tabs principales
    tabs = ft.Tabs(
        selected_index=0,
        animation_duration=300,
        tabs=[
            ft.Tab(
                text="⏰ Acciones Pendientes",
                content=create_general_tab(),
                icon=ft.Icons.PENDING_ACTIONS
            ),
            ft.Tab(
                text="👥 Por Responsable",
                content=create_user_metrics_tab(),
                icon=ft.Icons.PEOPLE
            ),
            ft.Tab(
                text="🔬 Análisis Avanzado",
                content=create_analytics_tab(),
                icon=ft.Icons.ANALYTICS
            )
        ],
        expand=1,
        label_color=COLORS['text_secondary'],
        indicator_color=COLORS['accent']
    )
    
    page.add(tabs)

if __name__ == "__main__":
    ft.app(target=main)