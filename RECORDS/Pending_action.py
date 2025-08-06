import flet as ft
import pandas as pd
import sqlite3
from datetime import datetime, timedelta

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

class WOActionsDashboard:
    def __init__(self):
        # Usar la misma ruta de base de datos
        self.db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        # Ruta alternativa para desarrollo
        # self.db_path = r"C:\Users\J.Vazquez\Desktop\R4Database.db"
        
        self.df_original = None
        self.df_enriched = None
        
    def load_and_process_data(self):
        """Cargar y procesar datos con cruces a WOInquiry y Work_order_Actions_responsibles"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            # Query principal con cruces a WOInquiry y Work_order_Actions_responsibles
            query = """
            SELECT 
                am.*,
                wi.ItemNumber,
                wi.Description as WO_Description,
                wi.DueDt,
                wi.CreateDt,
                wi.Entity,
                wi.ProjectNo,
                wi.Sub,
                wi.St as WO_Status,
                wi.WoType,
                woar.User as ResponsibleUser,
                woar.email as ResponsibleEmail,
                woar.Name as ResponsibleName
            FROM ActionMessages am
            LEFT JOIN WOInquiry wi ON am.ItemNo = wi.WONo
            LEFT JOIN Work_order_Actions_responsibles woar ON am.OH = woar.User
            WHERE am.Type = 'WO' 
            AND (am.ACD = 'CN' OR am.ACD = 'AD')
            ORDER BY am.REQDT DESC, am.ItemNo
            """
            
            self.df_original = pd.read_sql_query(query, conn)
            conn.close()
            
            if len(self.df_original) == 0:
                print("⚠️ No se encontraron registros que cumplan los criterios")
                return False
            
            # Procesar datos
            self.process_enriched_data()
            
            print(f"✅ Datos procesados: {len(self.df_original)} registros con cruces completados")
            return True
            
        except Exception as e:
            print(f"❌ Error cargando y procesando datos: {e}")
            return False
    
    def process_enriched_data(self):
        """Procesar datos enriquecidos con información de WO y responsables"""
        df = self.df_original.copy()
        
        # Convertir columnas de fecha
        df['REQDT'] = pd.to_datetime(df['REQDT'], errors='coerce')
        df['ADT'] = pd.to_datetime(df['ADT'], errors='coerce')
        df['DueDt'] = pd.to_datetime(df['DueDt'], errors='coerce')
        df['CreateDt'] = pd.to_datetime(df['CreateDt'], errors='coerce')
        
        # Convertir columnas numéricas
        df['ActQty'] = pd.to_numeric(df['ActQty'], errors='coerce').fillna(0)
        df['FmQty'] = pd.to_numeric(df['FmQty'], errors='coerce').fillna(0)
        df['ToQty'] = pd.to_numeric(df['ToQty'], errors='coerce').fillna(0)
        df['StdCost'] = pd.to_numeric(df['StdCost'], errors='coerce').fillna(0)
        
        # Calcular métricas derivadas
        df['Valor_Action'] = df['ActQty'] * df['StdCost']
        df['Days_Since_Action'] = (datetime.now() - df['ADT']).dt.days
        df['Days_Until_Due'] = (df['DueDt'] - datetime.now()).dt.days
        
        # Limpiar datos de responsables
        df['ResponsibleUser'] = df['ResponsibleUser'].fillna('Sin Asignar')
        df['ResponsibleEmail'] = df['ResponsibleEmail'].fillna('No Email')
        df['ResponsibleName'] = df['ResponsibleName'].fillna('Sin Nombre')
        
        # Determinar estado del responsable
        df['ResponsibleStatus'] = df.apply(self._determine_responsible_status, axis=1)
        
        # Determinar urgencia basada en fecha de vencimiento
        df['Urgency'] = df['Days_Until_Due'].apply(self._determine_urgency)
        
        self.df_enriched = df
    
    def _determine_responsible_status(self, row):
        """Determinar estado del responsable"""
        if pd.notna(row['ResponsibleUser']) and row['ResponsibleUser'] != 'Sin Asignar':
            if pd.notna(row['ResponsibleEmail']) and row['ResponsibleEmail'] != 'No Email':
                return 'Completo'
            else:
                return 'Sin Email'
        else:
            return 'Sin Asignar'
    
    def _determine_urgency(self, days_until_due):
        """Determinar urgencia basada en días hasta vencimiento"""
        if pd.isna(days_until_due):
            return 'Sin Fecha'
        
        if days_until_due < 0:
            return 'Vencido'
        elif days_until_due <= 7:
            return 'Crítico'
        elif days_until_due <= 30:
            return 'Urgente'
        else:
            return 'Normal'
    
    def get_summary_metrics(self):
        """Calcular métricas resumen"""
        if self.df_enriched is None or len(self.df_enriched) == 0:
            return {
                'total_actions': 0,
                'with_responsible': 0,
                'without_responsible': 0,
                'with_email': 0,
                'cn_actions': 0,
                'ad_actions': 0,
                'overdue': 0,
                'critical': 0,
                'urgent': 0,
                'normal': 0,
                'unique_wos': 0,
                'unique_responsibles': 0,
                'total_value': 0
            }
        
        df = self.df_enriched
        
        # Métricas básicas
        total_actions = len(df)
        unique_wos = df['ItemNo'].nunique()
        unique_responsibles = len(df[df['ResponsibleUser'] != 'Sin Asignar']['ResponsibleUser'].unique())
        total_value = df['Valor_Action'].sum()
        
        # Métricas de responsables
        with_responsible = len(df[df['ResponsibleStatus'] == 'Completo'])
        without_responsible = len(df[df['ResponsibleStatus'] == 'Sin Asignar'])
        with_email = len(df[df['ResponsibleStatus'] == 'Completo'])
        
        # Métricas por tipo de acción
        cn_actions = len(df[df['ACD'] == 'CN'])
        ad_actions = len(df[df['ACD'] == 'AD'])
        
        # Métricas de urgencia
        urgency_counts = df['Urgency'].value_counts()
        overdue = urgency_counts.get('Vencido', 0)
        critical = urgency_counts.get('Crítico', 0)
        urgent = urgency_counts.get('Urgente', 0)
        normal = urgency_counts.get('Normal', 0)
        
        return {
            'total_actions': total_actions,
            'with_responsible': with_responsible,
            'without_responsible': without_responsible,
            'with_email': with_email,
            'cn_actions': cn_actions,
            'ad_actions': ad_actions,
            'overdue': overdue,
            'critical': critical,
            'urgent': urgent,
            'normal': normal,
            'unique_wos': unique_wos,
            'unique_responsibles': unique_responsibles,
            'total_value': total_value
        }
    
    def get_user_kpis(self):
        """Generar KPIs detallados por usuario"""
        if self.df_enriched is None or len(self.df_enriched) == 0:
            return pd.DataFrame()
        
        # Filtrar solo registros con usuario asignado
        df_with_users = self.df_enriched[
            (self.df_enriched['ResponsibleUser'].notna()) & 
            (self.df_enriched['ResponsibleUser'] != 'Sin Asignar')
        ].copy()
        
        if len(df_with_users) == 0:
            return pd.DataFrame()
        
        # Agrupar por usuario y calcular KPIs
        user_kpis = df_with_users.groupby(['ResponsibleUser', 'ResponsibleName', 'ResponsibleEmail']).agg({
            'id': 'count',  # Total de acciones
            'Valor_Action': 'sum',  # Valor total
            'ACD': lambda x: (x == 'CN').sum(),  # Acciones CN
            'Urgency': [
                lambda x: (x == 'Vencido').sum(),     # Vencidas
                lambda x: (x == 'Crítico').sum(),     # Críticas
                lambda x: (x == 'Urgente').sum(),     # Urgentes
                lambda x: (x == 'Normal').sum()       # Normales
            ],
            'Days_Since_Action': 'mean',  # Promedio días
            'ItemNo': 'nunique'  # WOs únicas
        }).reset_index()
        
        # Aplanar columnas multinivel
        user_kpis.columns = [
            'ResponsibleUser', 'ResponsibleName', 'ResponsibleEmail',
            'Total_Actions', 'Total_Value', 'CN_Actions',
            'Vencidas', 'Criticas', 'Urgentes', 'Normales',
            'Avg_Days', 'Unique_WOs'
        ]
        
        # Calcular métricas adicionales
        user_kpis['AD_Actions'] = user_kpis['Total_Actions'] - user_kpis['CN_Actions']
        user_kpis['Priority_Actions'] = user_kpis['Vencidas'] + user_kpis['Criticas']
        user_kpis['Completion_Rate'] = ((user_kpis['Normales'] / user_kpis['Total_Actions']) * 100).round(1)
        user_kpis['Risk_Score'] = (
            (user_kpis['Vencidas'] * 4) + 
            (user_kpis['Criticas'] * 3) + 
            (user_kpis['Urgentes'] * 2) + 
            (user_kpis['Normales'] * 1)
        )
        
        # Ordenar por total de acciones
        user_kpis = user_kpis.sort_values('Total_Actions', ascending=False).reset_index(drop=True)
        user_kpis['Ranking'] = range(1, len(user_kpis) + 1)
        
        return user_kpis

def create_metric_card(title, value, subtitle="", color=COLORS['accent']):
    """Crear tarjeta de métrica estilo ejecutivo"""
    return ft.Card(
        content=ft.Container(
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12,
            content=ft.Column([
                ft.Text(title, size=14, color=COLORS['text_secondary'], weight=ft.FontWeight.W_500),
                ft.Text(str(value), size=32, color=color, weight=ft.FontWeight.BOLD),
                ft.Text(subtitle, size=12, color=COLORS['text_secondary']) if subtitle else ft.Container()
            ], spacing=5)
        ),
        elevation=8
    )

def main(page: ft.Page):
    page.title = "WO Actions & Responsibles Dashboard"
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = COLORS['surface']
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20
    
    dashboard = WOActionsDashboard()
    
    def create_main_tab():
        main_container = ft.Column(spacing=20)
        status_text = ft.Text("Iniciando WO Actions dashboard...", color=COLORS['text_secondary'])
        
        # Tabla de detalles enriquecida
        data_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("WO Number", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Item Number", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Action", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Req Date", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Due Date", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Urgency", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Responsible", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Email", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Status", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Act Qty", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Value", color=COLORS['text_primary'])),
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=8,
        )
        
        def update_dashboard():
            main_container.controls.clear()
            data_table.rows.clear()
            
            status_text.value = "🔄 Cargando y procesando datos..."
            status_text.color = COLORS['accent']
            page.update()
            
            if dashboard.load_and_process_data():
                metrics = dashboard.get_summary_metrics()
                
                if metrics['total_actions'] == 0:
                    main_container.controls.append(
                        ft.Container(
                            content=ft.Column([
                                ft.Text("ℹ️ No hay WO Actions", 
                                       size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                ft.Text("No se encontraron registros que cumplan los criterios de filtrado", 
                                       size=16, color=COLORS['text_secondary']),
                                ft.ElevatedButton(
                                    "🔄 Actualizar Datos",
                                    on_click=lambda _: update_dashboard(),
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
                            ft.Text("🏭 WO Actions & Responsibles Dashboard", 
                                size=32, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            ft.Text(f"Análisis de {metrics['total_actions']:,} acciones con responsables asignados", 
                                size=16, color=COLORS['text_secondary']),
                            ft.Text("ActionMessages + WOInquiry + Work_order_Actions_responsibles", 
                                size=12, color=COLORS['warning'])
                        ], expand=True),
                        ft.Column([
                            ft.Row([
                                ft.ElevatedButton(
                                    "🔄 Actualizar Datos",
                                    on_click=lambda _: update_dashboard(),
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
                
                # Métricas principales
                main_metrics_row = ft.Row([
                    create_metric_card(
                        "Total WO Actions",
                        f"{metrics['total_actions']:,}",
                        f"Valor: ${metrics['total_value']:,.0f}",
                        COLORS['accent']
                    ),
                    create_metric_card(
                        "WOs Únicas",
                        f"{metrics['unique_wos']:,}",
                        "Work Orders diferentes",
                        COLORS['success']
                    ),
                    create_metric_card(
                        "Responsables",
                        f"{metrics['unique_responsibles']:,}",
                        "Usuarios únicos asignados",
                        COLORS['warning']
                    ),
                    create_metric_card(
                        "CN Actions",
                        f"{metrics['cn_actions']:,}",
                        f"AD Actions: {metrics['ad_actions']:,}",
                        COLORS['error']
                    ),
                ], wrap=True, spacing=15)
                
                # Métricas de responsables
                responsible_metrics_row = ft.Row([
                    create_metric_card(
                        "Con Responsable",
                        f"{metrics['with_responsible']:,}",
                        f"{(metrics['with_responsible']/metrics['total_actions']*100):.1f}% del total",
                        COLORS['success']
                    ),
                    create_metric_card(
                        "Sin Responsable",
                        f"{metrics['without_responsible']:,}",
                        "Requieren asignación",
                        COLORS['error']
                    ),
                    create_metric_card(
                        "Con Email",
                        f"{metrics['with_email']:,}",
                        "Contacto disponible",
                        COLORS['success']
                    ),
                    create_metric_card(
                        "Coverage",
                        f"{(metrics['with_responsible']/metrics['total_actions']*100):.1f}%",
                        "Cobertura de responsables",
                        COLORS['accent']
                    ),
                ], wrap=True, spacing=15)
                
                # Métricas de urgencia
                urgency_metrics_row = ft.Row([
                    create_metric_card(
                        "Vencidas",
                        f"{metrics['overdue']:,}",
                        "Pasaron fecha límite",
                        COLORS['error']
                    ),
                    create_metric_card(
                        "Críticas",
                        f"{metrics['critical']:,}",
                        "≤ 7 días",
                        COLORS['error']
                    ),
                    create_metric_card(
                        "Urgentes",
                        f"{metrics['urgent']:,}",
                        "8-30 días",
                        COLORS['warning']
                    ),
                    create_metric_card(
                        "Normales",
                        f"{metrics['normal']:,}",
                        "> 30 días",
                        COLORS['success']
                    ),
                ], wrap=True, spacing=15)
                
                # Llenar tabla con datos enriquecidos
                for _, row in dashboard.df_enriched.head(100).iterrows():
                    # Determinar colores según urgencia
                    if row['Urgency'] == 'Vencido':
                        urgency_color = COLORS['error']
                    elif row['Urgency'] == 'Crítico':
                        urgency_color = COLORS['error']
                    elif row['Urgency'] == 'Urgente':
                        urgency_color = COLORS['warning']
                    else:
                        urgency_color = COLORS['success']
                    
                    # Color del estado del responsable
                    if row['ResponsibleStatus'] == 'Completo':
                        status_color = COLORS['success']
                    elif row['ResponsibleStatus'] == 'Sin Email':
                        status_color = COLORS['warning']
                    else:
                        status_color = COLORS['error']
                    
                    data_table.rows.append(
                        ft.DataRow(
                            cells=[
                                ft.DataCell(ft.Text(str(row['ItemNo']), color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(str(row['ItemNumber']) if pd.notna(row['ItemNumber']) else "-", 
                                          color=COLORS['accent'], selectable=True)),
                                ft.DataCell(ft.Text(str(row['ACD']), color=COLORS['accent'], selectable=True)),
                                ft.DataCell(ft.Text(str(row['REQDT'])[:10] if pd.notna(row['REQDT']) else "-", 
                                          color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(str(row['DueDt'])[:10] if pd.notna(row['DueDt']) else "-", 
                                          color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(row['Urgency'], color=urgency_color, selectable=True)),
                                ft.DataCell(ft.Text(str(row['ResponsibleUser'])[:20] + "..." if len(str(row['ResponsibleUser'])) > 20 else str(row['ResponsibleUser']), 
                                          color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(str(row['ResponsibleEmail'])[:25] + "..." if len(str(row['ResponsibleEmail'])) > 25 else str(row['ResponsibleEmail']), 
                                          color=COLORS['text_secondary'], selectable=True)),
                                ft.DataCell(ft.Text(row['ResponsibleStatus'], color=status_color, selectable=True)),
                                ft.DataCell(ft.Text(f"{row['ActQty']:,.0f}", color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(f"${row['Valor_Action']:,.2f}", color=COLORS['success'], selectable=True)),
                            ]
                        )
                    )
                
                # Contenedor de tabla
                table_container = ft.Container(
                    content=ft.Column([
                        ft.Text("📋 WO Actions con Responsables (Top 100)", 
                               size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                        ft.Text("Datos enriquecidos con información de WO y responsables - Celdas seleccionables", 
                               size=12, color=COLORS['text_secondary']),
                        ft.Container(
                            content=ft.Column([data_table], scroll=ft.ScrollMode.ALWAYS),
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
                
                # Agregar todos los componentes
                main_container.controls.extend([
                    header,
                    ft.Container(
                        content=ft.Column([
                            ft.Text("📈 Métricas Generales", 
                                   size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            main_metrics_row,
                        ]),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=12
                    ),
                    ft.Container(
                        content=ft.Column([
                            ft.Text("👥 Asignación de Responsables", 
                                   size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            responsible_metrics_row,
                        ]),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=12
                    ),
                    ft.Container(
                        content=ft.Column([
                            ft.Text("⚡ Urgencia y Vencimientos", 
                                   size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            urgency_metrics_row,
                        ]),
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
        
        update_dashboard()
        return main_container
    
    def create_user_kpis_tab():
        """Crear tab específico para KPIs por usuario"""
        user_container = ft.Column(spacing=20)
        
        # Tabla de KPIs por usuario
        user_kpis_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Rank", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Usuario", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Nombre", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Email", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Total Actions", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("WOs Únicas", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("CN Actions", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("AD Actions", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Vencidas", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Críticas", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Valor Total", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Risk Score", color=COLORS['text_primary'])),
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=8,
        )
        
        def update_user_kpis():
            user_container.controls.clear()
            user_kpis_table.rows.clear()
            
            if dashboard.df_enriched is None:
                user_container.controls.append(
                    ft.Text("⚠️ Carga primero el dashboard principal", 
                           size=16, color=COLORS['warning'])
                )
                page.update()
                return
            
            user_kpis = dashboard.get_user_kpis()
            
            if len(user_kpis) == 0:
                user_container.controls.append(
                    ft.Container(
                        content=ft.Text("⚠️ No hay usuarios con acciones asignadas", 
                                       size=18, color=COLORS['warning']),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=8
                    )
                )
                page.update()
                return
            
            # Métricas resumen de usuarios
            total_users = len(user_kpis)
            users_with_overdue = len(user_kpis[user_kpis['Vencidas'] > 0])
            users_with_critical = len(user_kpis[user_kpis['Criticas'] > 0])
            avg_actions_per_user = user_kpis['Total_Actions'].mean()
            total_value_all_users = user_kpis['Total_Value'].sum()
            
            # Header para KPIs de usuarios
            user_header = ft.Container(
                content=ft.Column([
                    ft.Text("👥 KPIs por Usuario Responsable", 
                           size=28, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text(f"Análisis detallado de {total_users} usuarios con acciones asignadas", 
                           size=16, color=COLORS['text_secondary']),
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )
            
            # Métricas de usuarios
            user_metrics_row = ft.Row([
                create_metric_card(
                    "Total Usuarios",
                    f"{total_users:,}",
                    "Con acciones asignadas",
                    COLORS['accent']
                ),
                create_metric_card(
                    "Usuarios c/ Vencidas",
                    f"{users_with_overdue:,}",
                    f"{(users_with_overdue/total_users*100):.1f}% del total",
                    COLORS['error']
                ),
                create_metric_card(
                    "Usuarios c/ Críticas",
                    f"{users_with_critical:,}",
                    f"{(users_with_critical/total_users*100):.1f}% del total",
                    COLORS['warning']
                ),
                create_metric_card(
                    "Promedio Actions",
                    f"{avg_actions_per_user:.1f}",
                    f"Valor total: ${total_value_all_users:,.0f}",
                    COLORS['success']
                ),
            ], wrap=True, spacing=15)
            
            # Llenar tabla de KPIs por usuario
            for _, user in user_kpis.iterrows():
                # Determinar colores según ranking y riesgo
                if user['Ranking'] <= 3:
                    rank_color = COLORS['error']  # Top 3
                elif user['Ranking'] <= 10:
                    rank_color = COLORS['warning']  # Top 10
                else:
                    rank_color = COLORS['success']
                
                # Color del risk score
                if user['Risk_Score'] >= 20:
                    risk_color = COLORS['error']
                elif user['Risk_Score'] >= 10:
                    risk_color = COLORS['warning']
                else:
                    risk_color = COLORS['success']
                
                user_kpis_table.rows.append(
                    ft.DataRow(
                        cells=[
                            ft.DataCell(ft.Text(f"#{user['Ranking']}", color=rank_color, weight=ft.FontWeight.BOLD, selectable=True)),
                            ft.DataCell(ft.Text(str(user['ResponsibleUser']), color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(user['ResponsibleName'])[:20] + "..." if len(str(user['ResponsibleName'])) > 20 else str(user['ResponsibleName']), 
                                              color=COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(str(user['ResponsibleEmail'])[:25] + "..." if len(str(user['ResponsibleEmail'])) > 25 else str(user['ResponsibleEmail']), 
                                              color=COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Total_Actions']:,}", color=COLORS['accent'], weight=ft.FontWeight.BOLD, selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Unique_WOs']:,}", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['CN_Actions']:,}", color=COLORS['error'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['AD_Actions']:,}", color=COLORS['success'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Vencidas']:,}", color=COLORS['error'] if user['Vencidas'] > 0 else COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Criticas']:,}", color=COLORS['warning'] if user['Criticas'] > 0 else COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(f"${user['Total_Value']:,.0f}", color=COLORS['success'], selectable=True)),
                            ft.DataCell(ft.Text(f"{user['Risk_Score']:.0f}", color=risk_color, weight=ft.FontWeight.BOLD, selectable=True)),
                        ]
                    )
                )
            
            # Contenedor de tabla de usuarios
            user_table_container = ft.Container(
                content=ft.Column([
                    ft.Text("📊 Ranking de Usuarios por Acciones", 
                           size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Risk Score = (Vencidas×4) + (Críticas×3) + (Urgentes×2) + (Normales×1)", 
                           size=12, color=COLORS['text_secondary']),
                    ft.Container(
                        content=ft.Column([user_kpis_table], scroll=ft.ScrollMode.ALWAYS),
                        border=ft.border.all(1, COLORS['secondary']),
                        border_radius=8,
                        padding=15,
                        height=600,
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
                        ft.Text("📈 Métricas de Usuarios", 
                               size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                        user_metrics_row,
                    ]),
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12
                ),
                user_table_container
            ])
            
            page.update()
        
        update_user_kpis()
        return user_container
    
    # Crear tabs
    tabs = ft.Tabs(
        selected_index=0,
        animation_duration=300,
        tabs=[
            ft.Tab(
                text="🏭 WO Actions Overview",
                content=create_main_tab()
            ),
            ft.Tab(
                text="👥 KPIs por Usuario",
                content=create_user_kpis_tab()
            )
        ],
        expand=1
    )
    
    page.add(tabs)

if __name__ == "__main__":
    ft.app(target=main)