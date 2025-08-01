import flet as ft
import pandas as pd
import sqlite3
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Para evitar problemas de display
import io
import base64

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

class InventoryDashboard:
    def __init__(self):
        self.db_path = r"C:\Users\J.Vazquez\Desktop\R4Database.db"
        self.df_original = None
        self.df_costs = None
        self.df_filtered = None
        
    def load_data_from_db(self):
        """Cargar datos desde la base de datos SQLite"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            # Cargar tabla principal PR561
            query_pr561 = "SELECT * FROM pr561"
            self.df_original = pd.read_sql_query(query_pr561, conn)
            
            # Cargar tabla de costos IN92
            query_in92 = "SELECT ItemNo, STDCost FROM in92"
            self.df_costs = pd.read_sql_query(query_in92, conn)
            
            conn.close()
            print(f"✅ Datos cargados: {len(self.df_original)} registros totales")
            return True
        except Exception as e:
            print(f"Error cargando datos: {e}")
            return False
    
    def process_data(self):
        """Procesar datos aplicando la lógica de cobertura SECUENCIAL por componente"""
        if self.df_original is None or self.df_costs is None:
            return False
            
        df = self.df_original.copy()
        print(f"📊 Registros iniciales: {len(df)}")
        
        # Hacer el cruce con la tabla de costos
        df = df.merge(self.df_costs, left_on='Component', right_on='ItemNo', how='left')
        
        # Asegurar que 'ReqDate' es tipo fecha
        df['ReqDate'] = pd.to_datetime(df['ReqDate'], errors='coerce')
        
        # Convertir columnas numéricas
        df['QtyOh'] = pd.to_numeric(df['QtyOh'], errors='coerce').fillna(0)
        df['ReqQty'] = pd.to_numeric(df['ReqQty'], errors='coerce').fillna(0)
        df['QtyPending'] = pd.to_numeric(df['QtyPending'], errors='coerce').fillna(0)
        df['STDCost'] = pd.to_numeric(df['STDCost'], errors='coerce').fillna(0)
        
        # *** CONVERTIR SRT A MAYÚSCULAS PARA COMPARACIÓN ***
        df['Srt'] = df['Srt'].astype(str).str.upper()
        
        # *** DICCIONARIO DE SORT CODES A EXCLUIR ***
        sort_codes_excluidos = {
            'FS', 'FS1', 'FS2', 'FSR', 'NC', 'NC1', 'NC2', 'NCB', 
            'NCI', 'NCL', 'NCP', 'NCS', 'NCT', 'PH'
        }
        
        # Filtrar por sort codes ANTES de otros filtros
        df_sin_excluidos = df[~df['Srt'].isin(sort_codes_excluidos)].copy()
        
        # Fecha actual
        hoy = pd.Timestamp.today()
        
        # Filtrar past due o del mes de junio
        df_filtrado = df_sin_excluidos[
            (df_sin_excluidos['ReqDate'] < hoy) | 
            (df_sin_excluidos['ReqDate'].dt.month == 6)
        ].copy()
        
        # *** FILTRO PRINCIPAL: Solo considerar líneas con QtyPending > 0 ***
        df_filtrado = df_filtrado[df_filtrado['QtyPending'] > 0].copy()
        
        if len(df_filtrado) == 0:
            print("⚠️ No hay líneas pendientes después de aplicar filtros")
            self.df_filtered = df_filtrado
            return True
        
        # *** ORDENAR POR COMPONENT Y FECHA PARA COBERTURA SECUENCIAL ***
        df_filtrado.sort_values(by=['Component', 'ReqDate'], inplace=True)
        
        # Crear columnas nuevas
        df_filtrado['COB'] = ''
        df_filtrado['Balance'] = 0.0
        df_filtrado['Valor_No_Cubierto'] = 0.0
        df_filtrado['Valor_Cubierto'] = 0.0
        df_filtrado['Qty_Faltante'] = 0.0
        
        # *** COBERTURA SECUENCIAL POR COMPONENTE ***
        for component, grupo in df_filtrado.groupby('Component'):
            # Obtener inventario disponible para este componente
            inventario_disponible = float(grupo['QtyOh'].iloc[0])
            
            # Procesar demandas secuencialmente (ya ordenadas por fecha)
            for idx, row in grupo.iterrows():
                qty_pending = float(row['QtyPending'])
                std_cost = float(row['STDCost'])
                
                if inventario_disponible >= qty_pending:
                    # ✅ Tenemos suficiente inventario para cubrir COMPLETAMENTE esta demanda
                    df_filtrado.at[idx, 'COB'] = 'Sí'
                    df_filtrado.at[idx, 'Valor_Cubierto'] = qty_pending * std_cost
                    df_filtrado.at[idx, 'Valor_No_Cubierto'] = 0
                    df_filtrado.at[idx, 'Qty_Faltante'] = 0
                    
                    # Consumir inventario
                    inventario_disponible -= qty_pending
                    df_filtrado.at[idx, 'Balance'] = inventario_disponible
                    
                elif inventario_disponible > 0:
                    # 🟡 Tenemos inventario PARCIAL para esta demanda
                    df_filtrado.at[idx, 'COB'] = 'Parcial'
                    qty_cubierta = inventario_disponible
                    qty_faltante = qty_pending - inventario_disponible
                    
                    df_filtrado.at[idx, 'Valor_Cubierto'] = qty_cubierta * std_cost
                    df_filtrado.at[idx, 'Valor_No_Cubierto'] = qty_faltante * std_cost
                    df_filtrado.at[idx, 'Qty_Faltante'] = qty_faltante
                    
                    # Inventario se agota completamente
                    inventario_disponible = 0
                    df_filtrado.at[idx, 'Balance'] = 0
                    
                else:
                    # ❌ No tenemos inventario para esta demanda
                    df_filtrado.at[idx, 'COB'] = 'No'
                    df_filtrado.at[idx, 'Valor_Cubierto'] = 0
                    df_filtrado.at[idx, 'Valor_No_Cubierto'] = qty_pending * std_cost
                    df_filtrado.at[idx, 'Qty_Faltante'] = qty_pending
                    df_filtrado.at[idx, 'Balance'] = 0
        
        self.df_filtered = df_filtrado
        
        # *** GUARDAR WO CLEAR EN BASE DE DATOS ***
        wo_clear_result = self.save_wo_clear_to_db()
        if wo_clear_result["success"]:
            print(f"📊 {wo_clear_result['message']}")
        
        return True
    
    def get_summary_metrics(self):
        """Calcular métricas resumen para el dashboard"""
        if self.df_filtered is None or len(self.df_filtered) == 0:
            return {
                'total_lines': 0, 'covered_lines': 0, 'partial_lines': 0, 'not_covered_lines': 0,
                'coverage_percentage': 0, 'total_value_required': 0, 'total_value_not_covered': 0,
                'total_value_covered': 0, 'total_qty_pending': 0, 'wo_clear_count': 0,
                'wo_clear_value': 0, 'wo_total_count': 0, 'wo_clear_percentage': 0
            }
        
        total_lines = len(self.df_filtered)
        covered_lines = len(self.df_filtered[self.df_filtered['COB'] == 'Sí'])
        partial_lines = len(self.df_filtered[self.df_filtered['COB'] == 'Parcial'])
        not_covered_lines = len(self.df_filtered[self.df_filtered['COB'] == 'No'])
        
        total_value_required = (self.df_filtered['QtyPending'] * self.df_filtered['STDCost']).sum()
        total_value_not_covered = self.df_filtered['Valor_No_Cubierto'].sum()
        total_value_covered = self.df_filtered['Valor_Cubierto'].sum()
        total_qty_pending = self.df_filtered['QtyPending'].sum()
        
        coverage_percentage = (covered_lines / total_lines * 100) if total_lines > 0 else 0
        wo_analysis = self._analyze_wo_coverage()
        
        return {
            'total_lines': total_lines, 'covered_lines': covered_lines, 'partial_lines': partial_lines,
            'not_covered_lines': not_covered_lines, 'coverage_percentage': coverage_percentage,
            'total_value_required': total_value_required, 'total_value_not_covered': total_value_not_covered,
            'total_value_covered': total_value_covered, 'total_qty_pending': total_qty_pending,
            'wo_clear_count': wo_analysis['wo_clear_count'], 'wo_clear_value': wo_analysis['wo_clear_value'],
            'wo_total_count': wo_analysis['wo_total_count'], 'wo_clear_percentage': wo_analysis['wo_clear_percentage']
        }
    
    def _analyze_wo_coverage(self):
        """Analizar cobertura por Work Order"""
        if self.df_filtered is None or len(self.df_filtered) == 0:
            return {'wo_clear_count': 0, 'wo_clear_value': 0, 'wo_total_count': 0, 'wo_clear_percentage': 0}
        
        wo_groups = self.df_filtered.groupby('WoNo')
        wo_clear_count = 0
        wo_clear_value = 0
        wo_total_count = len(wo_groups)
        
        for wo_number, wo_data in wo_groups:
            all_covered = all(wo_data['COB'] == 'Sí')
            if all_covered:
                wo_clear_count += 1
                wo_value = (wo_data['QtyPending'] * wo_data['STDCost']).sum()
                wo_clear_value += wo_value
        
        wo_clear_percentage = (wo_clear_count / wo_total_count * 100) if wo_total_count > 0 else 0
        
        return {
            'wo_clear_count': wo_clear_count, 'wo_clear_value': wo_clear_value,
            'wo_total_count': wo_total_count, 'wo_clear_percentage': wo_clear_percentage
        }
    
    def save_wo_clear_to_db(self):
        """Guardar Work Orders completamente cubiertas en tabla WOClear"""
        if self.df_filtered is None or len(self.df_filtered) == 0:
            return {"success": False, "message": "No hay datos para procesar"}
        
        try:
            conn = sqlite3.connect(self.db_path)
            conn.execute("DROP TABLE IF EXISTS WOClear")
            
            create_table_sql = """
            CREATE TABLE WOClear (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                WoNo TEXT, Component TEXT, ReqDate TEXT, QtyOh REAL, QtyPending REAL,
                STDCost REAL, Valor_Cubierto REAL, Valor_No_Cubierto REAL, Balance REAL,
                Srt TEXT, Project TEXT, Entity TEXT, Description TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
            conn.execute(create_table_sql)
            
            wo_clear_data = []
            for wo_number, wo_data in self.df_filtered.groupby('WoNo'):
                all_covered = all(wo_data['COB'] == 'Sí')
                if all_covered:
                    for _, row in wo_data.iterrows():
                        wo_clear_data.append({
                            'WoNo': row['WoNo'], 'Component': row['Component'],
                            'ReqDate': str(row['ReqDate'])[:10], 'QtyOh': float(row['QtyOh']),
                            'QtyPending': float(row['QtyPending']), 'STDCost': float(row['STDCost']),
                            'Valor_Cubierto': float(row['Valor_Cubierto']), 'Valor_No_Cubierto': float(row['Valor_No_Cubierto']),
                            'Balance': float(row['Balance']), 'Srt': str(row['Srt']),
                            'Project': str(row.get('Project', '')), 'Entity': str(row.get('Entity', '')),
                            'Description': str(row.get('Description', ''))
                        })
            
            if wo_clear_data:
                df_wo_clear = pd.DataFrame(wo_clear_data)
                df_wo_clear.to_sql('WOClear', conn, if_exists='append', index=False)
                conn.close()
                return {"success": True, "message": f"Tabla WOClear actualizada con {len(wo_clear_data)} registros"}
            else:
                conn.close()
                return {"success": False, "message": "No hay Work Orders completamente cubiertas"}
                
        except Exception as e:
            return {"success": False, "message": f"Error: {str(e)}"}
    
    def load_wo_clear_from_db(self):
        """Cargar datos de la tabla WOClear"""
        try:
            conn = sqlite3.connect(self.db_path)
            query = "SELECT * FROM WOClear ORDER BY WoNo, Component"
            df_wo_clear = pd.read_sql_query(query, conn)
            conn.close()
            return df_wo_clear
        except Exception as e:
            return pd.DataFrame()
    
    def get_wo_clear_filters(self):
        """Obtener valores únicos para filtros de WO Clear"""
        try:
            conn = sqlite3.connect(self.db_path)
            wo_query = "SELECT DISTINCT WoNo FROM WOClear ORDER BY WoNo"
            wo_numbers = pd.read_sql_query(wo_query, conn)['WoNo'].tolist()
            
            project_query = "SELECT DISTINCT Project FROM WOClear WHERE Project != '' ORDER BY Project"
            projects = pd.read_sql_query(project_query, conn)['Project'].tolist()
            
            entity_query = "SELECT DISTINCT Entity FROM WOClear WHERE Entity != '' ORDER BY Entity"
            entities = pd.read_sql_query(entity_query, conn)['Entity'].tolist()
            
            conn.close()
            return {'wo_numbers': wo_numbers, 'projects': projects, 'entities': entities}
        except Exception as e:
            return {'wo_numbers': [], 'projects': [], 'entities': []}
    
    def get_wo_analysis(self, wo_number=None):
        """Analizar una WO específica o todas las WOs para Pareto"""
        if self.df_original is None:
            return None
        
        try:
            conn = sqlite3.connect(self.db_path)
            
            if wo_number:
                query = """
                SELECT pr.*, in92.STDCost 
                FROM pr561 pr 
                LEFT JOIN in92 ON pr.Component = in92.ItemNo 
                WHERE pr.WoNo = ? 
                ORDER BY pr.ReqDate, pr.Component
                """
                df_wo = pd.read_sql_query(query, conn, params=[wo_number])
            else:
                query = """
                SELECT pr.WoNo, pr.Component, pr.QtyPending, pr.ReqQty, in92.STDCost,
                       pr.Project, pr.Entity
                FROM pr561 pr 
                LEFT JOIN in92 ON pr.Component = in92.ItemNo 
                WHERE pr.QtyPending > 0
                ORDER BY pr.WoNo, pr.Component
                """
                df_wo = pd.read_sql_query(query, conn)
            
            conn.close()
            
            if len(df_wo) == 0:
                return None
            
            df_wo['QtyPending'] = pd.to_numeric(df_wo['QtyPending'], errors='coerce').fillna(0)
            df_wo['ReqQty'] = pd.to_numeric(df_wo['ReqQty'], errors='coerce').fillna(0)
            df_wo['STDCost'] = pd.to_numeric(df_wo['STDCost'], errors='coerce').fillna(0)
            
            df_wo['Valor_Pendiente'] = df_wo['QtyPending'] * df_wo['STDCost']
            df_wo['Valor_Total'] = df_wo['ReqQty'] * df_wo['STDCost']
            df_wo['Status'] = df_wo['QtyPending'].apply(lambda x: 'Surtido' if x == 0 else 'Pendiente')
            
            return df_wo
            
        except Exception as e:
            return None
    
    def get_pareto_data(self):
        """Generar datos para gráfico de Pareto de WOs más costosas"""
        df_all = self.get_wo_analysis()
        
        if df_all is None or len(df_all) == 0:
            return None
        
        wo_summary = df_all.groupby('WoNo').agg({
            'Valor_Pendiente': 'sum', 'Valor_Total': 'sum', 'QtyPending': 'sum',
            'Project': 'first', 'Entity': 'first'
        }).reset_index()
        
        wo_summary = wo_summary.sort_values('Valor_Pendiente', ascending=False)
        wo_summary['Pct_Individual'] = (wo_summary['Valor_Pendiente'] / wo_summary['Valor_Pendiente'].sum()) * 100
        wo_summary['Pct_Acumulado'] = wo_summary['Pct_Individual'].cumsum()
        wo_summary['Ranking'] = range(1, len(wo_summary) + 1)
        
        return wo_summary.head(20)
    
    def get_wo_clear_pareto_data(self):
        """Generar datos de Pareto específicamente para WO Clear"""
        try:
            conn = sqlite3.connect(self.db_path)
            query = """
            SELECT WoNo, SUM(Valor_Cubierto) as Total_Value, 
                   COUNT(*) as Component_Count,
                   MAX(Project) as Project, MAX(Entity) as Entity
            FROM WOClear 
            GROUP BY WoNo 
            ORDER BY Total_Value DESC
            """
            df_wo_clear = pd.read_sql_query(query, conn)
            conn.close()
            
            if len(df_wo_clear) == 0:
                return None
            
            # Calcular porcentajes
            total_value = df_wo_clear['Total_Value'].sum()
            df_wo_clear['Pct_Individual'] = (df_wo_clear['Total_Value'] / total_value) * 100
            df_wo_clear['Pct_Acumulado'] = df_wo_clear['Pct_Individual'].cumsum()
            df_wo_clear['Ranking'] = range(1, len(df_wo_clear) + 1)
            
            return df_wo_clear.head(20)
            
        except Exception as e:
            print(f"Error generando datos Pareto WO Clear: {e}")
            return None
    
    def create_pareto_chart(self, pareto_data):
        """Crear gráfico de Pareto y retornar como base64"""
        if pareto_data is None or len(pareto_data) == 0:
            return None
        
        try:
            # Configurar estilo oscuro
            plt.style.use('dark_background')
            fig, ax1 = plt.subplots(figsize=(16, 8))
            fig.patch.set_facecolor('#0f172a')
            ax1.set_facecolor('#1e293b')
            
            # Colores del tema
            bar_color = '#0ea5e9'  # Sky 500
            line_color = '#10b981'  # Emerald 500
            text_color = '#f8fafc'  # Slate 50
            grid_color = '#334155'  # Slate 700
            
            # Gráfico de barras para valores
            x_pos = range(len(pareto_data))
            bars = ax1.bar(x_pos, pareto_data['Total_Value'], 
                          color=bar_color, alpha=0.8, edgecolor='white', linewidth=0.5)
            
            ax1.set_xlabel('Work Orders (Top 20)', fontsize=12, color=text_color, fontweight='bold')
            ax1.set_ylabel('Valor Total ($)', fontsize=12, color=bar_color, fontweight='bold')
            ax1.tick_params(axis='y', labelcolor=bar_color, colors=bar_color)
            ax1.tick_params(axis='x', labelcolor=text_color, colors=text_color, rotation=45)
            
            # Formatear etiquetas del eje X (WO Numbers)
            wo_labels = [str(wo)[:8] + '...' if len(str(wo)) > 8 else str(wo) for wo in pareto_data['WoNo']]
            ax1.set_xticks(x_pos)
            ax1.set_xticklabels(wo_labels)
            
            # Formatear eje Y con formato de moneda
            ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
            
            # Crear segundo eje Y para porcentaje acumulado
            ax2 = ax1.twinx()
            line = ax2.plot(x_pos, pareto_data['Pct_Acumulado'], 
                           color=line_color, marker='o', linewidth=3, markersize=6, 
                           markerfacecolor='white', markeredgecolor=line_color, markeredgewidth=2)
            
            ax2.set_ylabel('Porcentaje Acumulado (%)', fontsize=12, color=line_color, fontweight='bold')
            ax2.tick_params(axis='y', labelcolor=line_color, colors=line_color)
            ax2.set_ylim(0, 100)
            
            # Línea del 80% (Principio de Pareto)
            ax2.axhline(y=80, color='#ef4444', linestyle='--', linewidth=2, alpha=0.8)
            ax2.text(len(pareto_data)/2, 82, '80% Rule', ha='center', va='bottom', 
                    color='#ef4444', fontsize=10, fontweight='bold')
            
            # Agregar valores en las barras
            for i, (bar, value) in enumerate(zip(bars, pareto_data['Total_Value'])):
                if value > 0:
                    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(pareto_data['Total_Value'])*0.01,
                            f'${value:,.0f}', ha='center', va='bottom', color=text_color, 
                            fontsize=8, fontweight='bold', rotation=90)
            
            # Agregar valores de porcentaje en la línea
            for i, (x, pct) in enumerate(zip(x_pos, pareto_data['Pct_Acumulado'])):
                if i % 2 == 0 or pct >= 80:  # Mostrar solo algunos valores para evitar sobrecarga
                    ax2.text(x, pct + 2, f'{pct:.1f}%', ha='center', va='bottom',
                            color=line_color, fontsize=8, fontweight='bold')
            
            # Título y subtítulo
            fig.suptitle('📊 Análisis de Pareto - Work Orders Clear por Valor', 
                        fontsize=16, color=text_color, fontweight='bold', y=0.95)
            ax1.text(len(pareto_data)/2, max(pareto_data['Total_Value'])*0.85, 
                    f'Top 20 WOs Clear • Total: ${pareto_data["Total_Value"].sum():,.0f}',
                    ha='center', va='center', color=text_color, fontsize=12,
                    bbox=dict(boxstyle="round,pad=0.3", facecolor='#334155', alpha=0.8))
            
            # Configurar grid
            ax1.grid(True, alpha=0.3, color=grid_color)
            ax2.grid(False)
            
            # Ajustar layout
            plt.tight_layout()
            
            # Convertir a base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight', 
                       facecolor='#0f172a', edgecolor='none')
            buffer.seek(0)
            img_base64 = base64.b64encode(buffer.getvalue()).decode()
            plt.close()
            
            return img_base64
            
        except Exception as e:
            print(f"Error creando gráfico Pareto: {e}")
            plt.close()
            return None

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

def create_export_card(dashboard, page):
    """Crear tarjeta de exportación con estilo ejecutivo"""
    export_button_ref = ft.Ref[ft.ElevatedButton]()
    status_text_ref = ft.Ref[ft.Text]()
    
    def handle_export(_):
        export_button_ref.current.text = "⏳ Exportando..."
        export_button_ref.current.disabled = True
        status_text_ref.current.value = "🔄 Procesando exportación..."
        export_button_ref.current.update()
        status_text_ref.current.update()
        
        try:
            # Lógica de exportación simplificada
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            export_path = f"Analisis_Cobertura_{timestamp}.xlsx"
            
            if dashboard.df_filtered is not None and len(dashboard.df_filtered) > 0:
                dashboard.df_filtered.to_excel(export_path, index=False)
                status_text_ref.current.value = f"✅ Exportado: {export_path}"
                status_text_ref.current.color = COLORS['success']
                export_button_ref.current.text = "✅ Completado"
            else:
                status_text_ref.current.value = "❌ No hay datos para exportar"
                status_text_ref.current.color = COLORS['error']
                export_button_ref.current.text = "❌ Sin datos"
                
        except Exception as e:
            status_text_ref.current.value = f"❌ Error: {str(e)}"
            status_text_ref.current.color = COLORS['error']
            export_button_ref.current.text = "❌ Error"
        
        status_text_ref.current.update()
        export_button_ref.current.update()
        
        # Restaurar botón después de 3 segundos
        import time
        import threading
        def restore_button():
            time.sleep(3)
            export_button_ref.current.text = "📊 Exportar a Excel"
            export_button_ref.current.disabled = False
            export_button_ref.current.update()
        
        threading.Thread(target=restore_button, daemon=True).start()
    
    return ft.Card(
        content=ft.Container(
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12,
            content=ft.Column([
                ft.Row([
                    ft.Text("📤 Exportar Análisis de Pendientes", 
                            size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                    ft.ElevatedButton("📊 Exportar a Excel", on_click=handle_export, bgcolor=COLORS['success'],
                                    color=COLORS['text_primary'], ref=export_button_ref, icon=ft.Icons.FILE_DOWNLOAD)
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Text("Listo para exportar líneas pendientes", size=12, color=COLORS['text_secondary'], ref=status_text_ref),
                ft.Divider(color=COLORS['secondary']),
                ft.Text("📋 Se exportarán:", size=12, color=COLORS['text_secondary']),
                ft.Column([
                    ft.Text("• Datos completos con cobertura calculada", size=11, color=COLORS['accent']),
                    ft.Text("• Columnas: Component, Dates, Quantities, Coverage", size=11, color=COLORS['text_secondary']),
                ], spacing=2)
            ], spacing=10)
        ),
        elevation=8
    )

def create_wo_clear_tab(dashboard):
    """Crear tab para visualizar Work Orders Clear"""
    
    # Estados para filtros
    selected_wo = ft.Ref[ft.Dropdown]()
    selected_project = ft.Ref[ft.Dropdown]()
    selected_entity = ft.Ref[ft.Dropdown]()
    wo_clear_table = ft.Ref[ft.DataTable]()
    wo_clear_metrics = ft.Ref[ft.Column]()
    pareto_image = ft.Ref[ft.Image]()
    pareto_container = ft.Ref[ft.Container]()
    
    def load_wo_clear_data():
        """Cargar datos de WO Clear con filtros"""
        try:
            df_wo_clear = dashboard.load_wo_clear_from_db()
            
            if len(df_wo_clear) == 0:
                wo_clear_metrics.current.controls = [
                    ft.Container(
                        content=ft.Text("⚠️ No hay datos en tabla WOClear. Ejecuta el Dashboard Principal primero.", 
                                       size=16, color=COLORS['warning']),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=8
                    )
                ]
                wo_clear_table.current.rows.clear()
                
                # Ocultar gráfico si no hay datos
                pareto_container.current.visible = False
                
                wo_clear_metrics.current.update()
                wo_clear_table.current.update()
                pareto_container.current.update()
                return
            
            # Aplicar filtros
            df_filtered = df_wo_clear.copy()
            
            if selected_wo.current and selected_wo.current.value and selected_wo.current.value != "Todas":
                df_filtered = df_filtered[df_filtered['WoNo'] == selected_wo.current.value]
            
            if selected_project.current and selected_project.current.value and selected_project.current.value != "Todos":
                df_filtered = df_filtered[df_filtered['Project'] == selected_project.current.value]
            
            if selected_entity.current and selected_entity.current.value and selected_entity.current.value != "Todas":
                df_filtered = df_filtered[df_filtered['Entity'] == selected_entity.current.value]
            
            # Actualizar métricas
            total_wos = len(df_filtered['WoNo'].unique()) if len(df_filtered) > 0 else 0
            total_components = len(df_filtered)
            total_value = df_filtered['Valor_Cubierto'].sum() if len(df_filtered) > 0 else 0
            
            wo_clear_metrics.current.controls = [
                ft.Row([
                    create_metric_card(
                        "WOs Clear Filtradas",
                        f"{total_wos:,}",
                        "Work Orders completamente cubiertas",
                        COLORS['success']
                    ),
                    create_metric_card(
                        "Componentes Totales",
                        f"{total_components:,}",
                        "Componentes en WOs filtradas",
                        COLORS['accent']
                    ),
                    create_metric_card(
                        "Valor Total",
                        f"${total_value:,.2f}",
                        "Valor de componentes cubiertos",
                        COLORS['success']
                    ),
                ], wrap=True, spacing=15)
            ]
            
            # Generar gráfico de Pareto
            pareto_data = dashboard.get_wo_clear_pareto_data()
            if pareto_data is not None and len(pareto_data) > 0:
                chart_image_b64 = dashboard.create_pareto_chart(pareto_data)
                if chart_image_b64:
                    pareto_image.current.src_base64 = chart_image_b64
                    pareto_container.current.visible = True
                    
                    # Agregar información del análisis
                    top_80_count = len(pareto_data[pareto_data['Pct_Acumulado'] <= 80])
                    if top_80_count == 0:
                        top_80_count = min(3, len(pareto_data))  # Al menos mostrar top 3
                    
                    top_80_value = pareto_data.head(top_80_count)['Total_Value'].sum()
                    total_pareto_value = pareto_data['Total_Value'].sum()
                    pareto_percentage = (top_80_value / total_pareto_value * 100) if total_pareto_value > 0 else 0
                    
                    # Actualizar texto del análisis
                    analysis_text = f"""📈 Análisis 80/20: Top {top_80_count} WOs representan {pareto_percentage:.1f}% del valor total
💰 Valor concentrado: ${top_80_value:,.0f} de ${total_pareto_value:,.0f} total
🎯 WOs críticas: {', '.join(pareto_data.head(3)['WoNo'].astype(str))}"""
                    
                    pareto_container.current.content.controls[1] = ft.Text(
                        analysis_text.strip(),
                        size=12, 
                        color=COLORS['text_secondary'],
                        text_align=ft.TextAlign.CENTER
                    )
                else:
                    pareto_container.current.visible = False
            else:
                pareto_container.current.visible = False
            
            # Actualizar tabla
            wo_clear_table.current.rows.clear()
            
            for _, row in df_filtered.head(50).iterrows():
                wo_clear_table.current.rows.append(
                    ft.DataRow(
                        cells=[
                            ft.DataCell(ft.Text(str(row['WoNo']), color=COLORS['text_primary'])),
                            ft.DataCell(ft.Text(str(row['Component']), color=COLORS['text_primary'])),
                            ft.DataCell(ft.Text(str(row['ReqDate']), color=COLORS['text_primary'])),
                            ft.DataCell(ft.Text(f"{row['QtyPending']:,.0f}", color=COLORS['warning'])),
                            ft.DataCell(ft.Text(f"{row['QtyOh']:,.0f}", color=COLORS['text_primary'])),
                            ft.DataCell(ft.Text(f"${row['STDCost']:,.2f}", color=COLORS['text_primary'])),
                            ft.DataCell(ft.Text(f"${row['Valor_Cubierto']:,.2f}", color=COLORS['success'])),
                            ft.DataCell(ft.Text(f"{row['Balance']:,.0f}", color=COLORS['accent'])),
                            ft.DataCell(ft.Text(str(row['Srt']), color=COLORS['text_secondary'])),
                        ]
                    )
                )
            
            wo_clear_metrics.current.update()
            wo_clear_table.current.update()
            pareto_container.current.update()
            pareto_image.current.update()
            
        except Exception as e:
            print(f"Error cargando WO Clear: {e}")
    
    def on_filter_change(e):
        """Manejar cambio de filtros"""
        load_wo_clear_data()
    
    def refresh_filters():
        """Actualizar opciones de filtros"""
        filters = dashboard.get_wo_clear_filters()
        
        # Actualizar dropdown de WOs
        wo_options = [ft.dropdown.Option("Todas")] + [ft.dropdown.Option(wo) for wo in filters['wo_numbers']]
        selected_wo.current.options = wo_options
        
        # Actualizar dropdown de Projects
        project_options = [ft.dropdown.Option("Todos")] + [ft.dropdown.Option(proj) for proj in filters['projects']]
        selected_project.current.options = project_options
        
        # Actualizar dropdown de Entities
        entity_options = [ft.dropdown.Option("Todas")] + [ft.dropdown.Option(ent) for ent in filters['entities']]
        selected_entity.current.options = entity_options
        
        selected_wo.current.update()
        selected_project.current.update()
        selected_entity.current.update()
    
    # Crear controles
    filters_row = ft.Row([
        ft.Column([
            ft.Text("Work Order:", size=12, color=COLORS['text_secondary']),
            ft.Dropdown(
                ref=selected_wo,
                options=[ft.dropdown.Option("Todas")],
                value="Todas",
                width=200,
                on_change=on_filter_change
            )
        ]),
        ft.Column([
            ft.Text("Project:", size=12, color=COLORS['text_secondary']),
            ft.Dropdown(
                ref=selected_project,
                options=[ft.dropdown.Option("Todos")],
                value="Todos",
                width=200,
                on_change=on_filter_change
            )
        ]),
        ft.Column([
            ft.Text("Entity:", size=12, color=COLORS['text_secondary']),
            ft.Dropdown(
                ref=selected_entity,
                options=[ft.dropdown.Option("Todas")],
                value="Todas",
                width=200,
                on_change=on_filter_change
            )
        ]),
        ft.ElevatedButton(
            "🔄 Actualizar",
            on_click=lambda _: (refresh_filters(), load_wo_clear_data()),
            bgcolor=COLORS['accent'],
            color=COLORS['text_primary']
        )
    ], spacing=20)
    
    # Contenedor del gráfico de Pareto
    pareto_chart_container = ft.Container(
        ref=pareto_container,
        visible=False,
        content=ft.Column([
            ft.Text("📊 Análisis de Pareto - WO Clear por Valor", 
                   size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
            ft.Text("Cargando análisis...", size=12, color=COLORS['text_secondary']),
            ft.Image(
                ref=pareto_image,
                width=1200,
                height=400,
                fit=ft.ImageFit.CONTAIN,
                border_radius=8
            ),
        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
        padding=ft.padding.all(20),
        bgcolor=COLORS['primary'],
        border_radius=12
    )
    
    # Tabla de WO Clear
    wo_clear_data_table = ft.DataTable(
        ref=wo_clear_table,
        columns=[
            ft.DataColumn(ft.Text("WO Number", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Component", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Req Date", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Qty Pending", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Qty OH", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("STD Cost", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Valor Cubierto", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Balance", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Sort Code", color=COLORS['text_primary'])),
        ],
        rows=[],
        bgcolor=COLORS['primary'],
        border_radius=8,
    )
    
    # Contenedor de métricas
    metrics_container = ft.Column(ref=wo_clear_metrics, controls=[])
    
    # Layout principal de la tab
    return ft.Column([
        ft.Container(
            content=ft.Column([
                ft.Text("🏭 Work Orders Completamente Cubiertas", 
                       size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Text("Solo WOs donde TODOS los componentes tienen cobertura completa", 
                       size=12, color=COLORS['text_secondary']),
                ft.Divider(color=COLORS['secondary']),
                filters_row,
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        ),
        
        # NUEVO: Gráfico de Pareto
        pareto_chart_container,
        
        ft.Container(
            content=metrics_container,
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        ),
        
        ft.Container(
            content=ft.Column([
                ft.Text("📋 Detalle de Componentes en WOs Clear", 
                       size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Container(
                    content=wo_clear_data_table,
                    border=ft.border.all(1, COLORS['secondary']),
                    border_radius=8,
                    padding=15,
                    height=400,
                    bgcolor=COLORS['primary'],
                )
            ], scroll=ft.ScrollMode.AUTO),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
    ], spacing=20, scroll=ft.ScrollMode.AUTO)

def main(page: ft.Page):
    page.title = "Dashboard Ejecutivo - Cobertura de Inventario"
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = COLORS['surface']
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20
    
    dashboard = InventoryDashboard()
    
    def create_main_tab():
        main_container = ft.Column(spacing=20)
        status_text = ft.Text("Iniciando dashboard...", color=COLORS['text_secondary'])
        
        # Tabla de detalles con tema dark
        data_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Component", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Req Date", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Sort Code", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Qty On Hand", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Qty Pending", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("STD Cost", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Coverage", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Valor Cubierto", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Valor No Cubierto", color=COLORS['text_primary'])),
                ft.DataColumn(ft.Text("Balance OH", color=COLORS['text_primary'])),
            ],
            rows=[],
            bgcolor=COLORS['primary'],
            border_radius=8,
        )
        
        def update_dashboard():
            main_container.controls.clear()
            data_table.rows.clear()
            
            if dashboard.load_data_from_db():
                if dashboard.process_data():
                    metrics = dashboard.get_summary_metrics()
                    
                    if metrics['total_lines'] == 0:
                        # No hay líneas pendientes
                        main_container.controls.append(
                            ft.Container(
                                content=ft.Column([
                                    ft.Text("ℹ️ No hay líneas pendientes", 
                                           size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                    ft.Text("Todas las líneas han sido surtidas completamente (QtyPending = 0)", 
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
                        status_text.value = "✅ Datos cargados - No hay líneas pendientes"
                        status_text.color = COLORS['success']
                        page.update()
                        return
                    
                    # Header ejecutivo
                    header = ft.Container(
                        content=ft.Row([
                            ft.Column([
                                ft.Text("📊 Dashboard Ejecutivo - Cobertura de Inventario", 
                                    size=32, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                ft.Text(f"Análisis de {metrics['total_lines']:,} líneas PENDIENTES de surtir", 
                                    size=16, color=COLORS['text_secondary']),
                                ft.Text("Filtrado: QtyPending > 0 | Excluye: FS, FS1, FS2, FSR, NC*, PH", 
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
                    
                    # Crear tarjeta de exportación
                    export_card = create_export_card(dashboard, page)
                    
                    # Métricas principales
                    metrics_row = ft.Row([
                        create_metric_card(
                            "Total Líneas Pendientes",
                            f"{metrics['total_lines']:,}",
                            f"Valor: ${metrics['total_value_required']:,.0f}",
                            COLORS['accent']
                        ),
                        create_metric_card(
                            "Cobertura Completa",
                            f"{metrics['covered_lines']:,}",
                            f"{metrics['coverage_percentage']:.1f}% | ${metrics['total_value_covered']:,.0f}",
                            COLORS['success']
                        ),
                        create_metric_card(
                            "Cobertura Parcial",
                            f"{metrics['partial_lines']:,}",
                            "Inventario parcial disponible",
                            COLORS['warning']
                        ),
                        create_metric_card(
                            "Sin Cobertura",
                            f"{metrics['not_covered_lines']:,}",
                            f"Valor: ${metrics['total_value_not_covered']:,.0f}",
                            COLORS['error']
                        ),
                    ], wrap=True, spacing=15)
                    
                    # Métricas de Work Orders
                    wo_metrics_row = ft.Row([
                        create_metric_card(
                            "WO Clear",
                            f"{metrics['wo_clear_count']:,}",
                            f"{metrics['wo_clear_percentage']:.1f}% de {metrics['wo_total_count']:,} WOs",
                            COLORS['success']
                        ),
                        create_metric_card(
                            "WO Clear Value",
                            f"${metrics['wo_clear_value']:,.2f}",
                            "Valor de WOs completamente cubiertas",
                            COLORS['success']
                        ),
                        create_metric_card(
                            "WOs Pendientes",
                            f"{metrics['wo_total_count'] - metrics['wo_clear_count']:,}",
                            "WOs con componentes faltantes",
                            COLORS['warning']
                        ),
                        create_metric_card(
                            "Total WOs",
                            f"{metrics['wo_total_count']:,}",
                            "Work Orders en análisis",
                            COLORS['accent']
                        ),
                    ], wrap=True, spacing=15)
                    
                    # Llenar tabla con datos filtrados
                    for _, row in dashboard.df_filtered.head(100).iterrows():
                        if row['COB'] == 'Sí':
                            coverage_color = COLORS['success']
                        elif row['COB'] == 'Parcial':
                            coverage_color = COLORS['warning']
                        else:
                            coverage_color = COLORS['error']
                        
                        data_table.rows.append(
                            ft.DataRow(
                                cells=[
                                    ft.DataCell(ft.Text(str(row['Component']), color=COLORS['text_primary'])),
                                    ft.DataCell(ft.Text(str(row['ReqDate'])[:10], color=COLORS['text_primary'])),
                                    ft.DataCell(ft.Text(str(row['Srt']), color=COLORS['text_secondary'])),
                                    ft.DataCell(ft.Text(f"{row['QtyOh']:,.0f}", color=COLORS['text_primary'])),
                                    ft.DataCell(ft.Text(f"{row['QtyPending']:,.0f}", color=COLORS['warning'])),
                                    ft.DataCell(ft.Text(f"${row['STDCost']:,.2f}", color=COLORS['text_primary'])),
                                    ft.DataCell(ft.Text(row['COB'], color=coverage_color)),
                                    ft.DataCell(ft.Text(f"${row['Valor_Cubierto']:,.2f}", color=COLORS['success'])),
                                    ft.DataCell(ft.Text(f"${row['Valor_No_Cubierto']:,.2f}", color=COLORS['error'])),
                                    ft.DataCell(ft.Text(f"{row['Balance']:,.0f}", color=COLORS['accent'])),
                                ]
                            )
                        )
                    
                    # Contenedor de tabla
                    table_container = ft.Container(
                        content=ft.Column([
                            ft.Text("📋 Detalle de Líneas PENDIENTES (Top 100)", 
                                   size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            ft.Text("Solo líneas con QtyPending > 0 - Material realmente por surtir", 
                                   size=12, color=COLORS['warning']),
                            ft.Container(
                                content=data_table,
                                border=ft.border.all(1, COLORS['secondary']),
                                border_radius=8,
                                padding=15,
                                height=400,
                                bgcolor=COLORS['primary'],
                            )
                        ], scroll=ft.ScrollMode.AUTO),
                        padding=ft.padding.all(20),
                        bgcolor=COLORS['primary'],
                        border_radius=12
                    )
                    
                    # Agregar todos los componentes
                    main_container.controls.extend([
                        header,
                        export_card,
                        ft.Container(
                            content=ft.Column([
                                ft.Text("📈 Métricas de Cobertura por Línea", 
                                       size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                metrics_row,
                            ]),
                            padding=ft.padding.all(20),
                            bgcolor=COLORS['primary'],
                            border_radius=12
                        ),
                        ft.Container(
                            content=ft.Column([
                                ft.Text("🏭 Métricas de Work Orders", 
                                       size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                                wo_metrics_row,
                            ]),
                            padding=ft.padding.all(20),
                            bgcolor=COLORS['primary'],
                            border_radius=12
                        ),
                        table_container
                    ])
                    
                    status_text.value = "✅ Datos cargados exitosamente"
                    status_text.color = COLORS['success']
                else:
                    status_text.value = "❌ Error procesando datos"
                    status_text.color = COLORS['error']
            else:
                status_text.value = "❌ Error conectando a la base de datos"
                status_text.color = COLORS['error']
            
            page.update()
        
        update_dashboard()
        return main_container
    
    # Crear tabs
    tabs = ft.Tabs(
        selected_index=0,
        animation_duration=300,
        tabs=[
            ft.Tab(
                text="📊 Dashboard Principal",
                content=create_main_tab()
            ),
            ft.Tab(
                text="🏭 WO Clear Status",
                content=create_wo_clear_tab(dashboard)
            )
        ],
        expand=1
    )
    
    page.add(tabs)

if __name__ == "__main__":
    ft.app(target=main)