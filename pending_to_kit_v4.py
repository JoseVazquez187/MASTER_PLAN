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
        # self.db_path = r"C:\Users\J.Vazquez\Desktop\R4Database.db"
        self.db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        
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
        
        # Hacer el cruce con la tabla de costos usando 'ItemNo' (que es la columna correcta)
        df = df.merge(self.df_costs, left_on='ItemNo', right_on='ItemNo', how='left')
        
        # Asegurar que 'ReqDate' es tipo fecha
        df['ReqDate'] = pd.to_datetime(df['ReqDate'], errors='coerce')
        
        # Convertir columnas numéricas usando los nombres correctos de las columnas
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
        
        # *** ORDENAR POR ITEMNO Y FECHA PARA COBERTURA SECUENCIAL ***
        df_filtrado.sort_values(by=['ItemNo', 'ReqDate'], inplace=True)
        
        # Crear columnas nuevas
        df_filtrado['COB'] = ''
        df_filtrado['Balance'] = 0.0
        df_filtrado['Valor_No_Cubierto'] = 0.0
        df_filtrado['Valor_Cubierto'] = 0.0
        df_filtrado['Qty_Faltante'] = 0.0
        
        # *** COBERTURA SECUENCIAL POR ITEMNO (que es el componente) ***
        for component, grupo in df_filtrado.groupby('ItemNo'):
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
        
        wo_groups = self.df_filtered.groupby('WONo')
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
                WONo TEXT, Component TEXT, ReqDate TEXT, QtyOh REAL, QtyPending REAL,
                STDCost REAL, Valor_Cubierto REAL, Valor_No_Cubierto REAL, Balance REAL,
                Srt TEXT, Project TEXT, Entity TEXT, Description TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
            conn.execute(create_table_sql)
            
            wo_clear_data = []
            for wo_number, wo_data in self.df_filtered.groupby('WONo'):
                all_covered = all(wo_data['COB'] == 'Sí')
                if all_covered:
                    for _, row in wo_data.iterrows():
                        wo_clear_data.append({
                            'WONo': row['WONo'], 'Component': row['ItemNo'],  # Usar ItemNo como Component
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
            query = "SELECT * FROM WOClear ORDER BY WONo, Component"
            df_wo_clear = pd.read_sql_query(query, conn)
            conn.close()
            return df_wo_clear
        except Exception as e:
            return pd.DataFrame()
    
    def get_wo_clear_filters(self):
        """Obtener valores únicos para filtros de WO Clear"""
        try:
            conn = sqlite3.connect(self.db_path)
            wo_query = "SELECT DISTINCT WONo FROM WOClear ORDER BY WONo"
            wo_numbers = pd.read_sql_query(wo_query, conn)['WONo'].tolist()
            
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
                LEFT JOIN in92 ON pr.ItemNo = in92.ItemNo 
                WHERE pr.WONo = ? 
                ORDER BY pr.ReqDate, pr.ItemNo
                """
                df_wo = pd.read_sql_query(query, conn, params=[wo_number])
            else:
                query = """
                SELECT pr.WONo, pr.ItemNo, pr.QtyPending, pr.ReqQty, in92.STDCost,
                       pr.Project, pr.Entity
                FROM pr561 pr 
                LEFT JOIN in92 ON pr.ItemNo = in92.ItemNo 
                WHERE pr.QtyPending > 0
                ORDER BY pr.WONo, pr.ItemNo
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
        
        wo_summary = df_all.groupby('WONo').agg({
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
            SELECT WONo, SUM(Valor_Cubierto) as Total_Value, 
                   COUNT(*) as Component_Count,
                   MAX(Project) as Project, MAX(Entity) as Entity
            FROM WOClear 
            GROUP BY WONo 
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
            wo_labels = [str(wo)[:8] + '...' if len(str(wo)) > 8 else str(wo) for wo in pareto_data['WONo']]
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

# FIN DE LA CLASE InventoryDashboard
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
            wo_labels = [str(wo)[:8] + '...' if len(str(wo)) > 8 else str(wo) for wo in pareto_data['WONo']]
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

# FIN DE LA CLASE InventoryDashboard
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
    """Crear tab para visualizar Work Orders Clear con tabla 80% y registro de WOs"""
    # Estados para filtros
    selected_wo = ft.Ref[ft.Dropdown]()
    selected_project = ft.Ref[ft.Dropdown]()
    selected_entity = ft.Ref[ft.Dropdown]()
    wo_clear_metrics = ft.Ref[ft.Column]()
    pareto_image = ft.Ref[ft.Image]()
    pareto_container = ft.Ref[ft.Container]()
    
    # NUEVOS: Referencias para tabla 80% y registro de WOs
    top_80_table = ft.Ref[ft.DataTable]()
    wo_entry = ft.Ref[ft.TextField]()
    register_status = ft.Ref[ft.Text]()
    register_button = ft.Ref[ft.ElevatedButton]()
    wo_notes = ft.Ref[ft.TextField]()  # NUEVO: Campo de notas
    download_button = ft.Ref[ft.ElevatedButton]()  # NUEVO: Botón de descarga
    
    # NUEVOS: Referencias para búsqueda de materiales WO
    wo_search_entry = ft.Ref[ft.TextField]()
    wo_search_button = ft.Ref[ft.ElevatedButton]()
    wo_search_status = ft.Ref[ft.Text]()
    wo_materials_table = ft.Ref[ft.DataTable]()
    wo_materials_container = ft.Ref[ft.Container]()
    
    # NUEVOS: Referencias para tabla WOInquiry
    wo_inquiry_table = ft.Ref[ft.DataTable]()
    wo_inquiry_container = ft.Ref[ft.Container]()
    wo_inquiry_status = ft.Ref[ft.Text]()
    
    def create_wo_clear_checked_table():
        """Crear tabla WOClear_checked si no existe"""
        try:
            conn = sqlite3.connect(dashboard.db_path)
            create_table_sql = """
            CREATE TABLE IF NOT EXISTS WOClear_checked (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                WONo TEXT UNIQUE,
                checked_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                user_action TEXT DEFAULT 'PENDING',
                notes TEXT
            )
            """
            conn.execute(create_table_sql)
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error creando tabla WOClear_checked: {e}")
            return False
    
    def search_wo_inquiry_data(wo_number):
        """Buscar WOs relacionadas en tabla WOInquiry CON OPERACIONES de Operation_WO"""
        try:
            conn = sqlite3.connect(dashboard.db_path)
            
            # Consultar WOInquiry buscando en WONo y ParentWO CON CRUCE A Operation_WO
            # COLUMNAS EXACTAS de Operation_WO
            query = """
            SELECT wi.Entity, wi.ProjectNo, wi.WONo, wi.SO_FCST, wi.Sub, wi.ParentWO, 
                wi.ItemNumber, wi.Rev, wi.Description, wi.AC, wi.DueDt, wi.CreateDt, 
                wi.WoType, wi.Srt, wi.PlanType, wi.St,
                ow.ItemNo, ow.Description as OpDescription, ow.Seq, ow.OpID, 
                ow.StartDate, ow.EndDate, ow.Status as OpStatus, ow.WrkCtr
            FROM WOInquiry wi
            LEFT JOIN Operation_WO ow ON wi.WONo = ow.WONo
            WHERE wi.WONo = ? OR wi.ParentWO = ?
            ORDER BY wi.WONo, wi.CreateDt, ow.Seq
            """
            
            df_wo_inquiry = pd.read_sql_query(query, conn, params=[wo_number, wo_number])
            conn.close()
            
            return df_wo_inquiry
            
        except Exception as e:
            print(f"Error consultando WOInquiry con Operation_WO: {e}")
            return pd.DataFrame()
    
    def update_wo_inquiry_table(df_wo_inquiry, searched_wo):
        """Actualizar tabla de WOInquiry CON INFORMACIÓN DE OPERACIONES"""
        wo_inquiry_table.current.rows.clear()
        
        if len(df_wo_inquiry) == 0:
            wo_inquiry_status.current.value = f"⚠️ No se encontraron WOs relacionadas para: {searched_wo}"
            wo_inquiry_status.current.color = COLORS['warning']
            wo_inquiry_container.current.visible = False
        else:
            for _, row in df_wo_inquiry.iterrows():
                # Determinar color del status usando St (Status)
                wo_status = str(row['St']) if pd.notna(row['St']) else "Sin Estado"
                if wo_status.upper() in ['COMPLETED', 'CLOSED', 'FINISHED', 'C']:
                    status_color = COLORS['success']
                elif wo_status.upper() in ['ACTIVE', 'OPEN', 'IN PROGRESS', 'A', 'O']:
                    status_color = COLORS['accent']
                else:
                    status_color = COLORS['warning']
                
                # Determinar color del status de operación
                op_status = str(row['OpStatus']) if pd.notna(row['OpStatus']) else "-"
                if op_status.upper() in ['COMPLETED', 'CLOSED', 'FINISHED', 'C']:
                    op_status_color = COLORS['success']
                elif op_status.upper() in ['ACTIVE', 'OPEN', 'IN PROGRESS', 'A', 'O']:
                    op_status_color = COLORS['accent']
                else:
                    op_status_color = COLORS['text_secondary']
                
                wo_inquiry_table.current.rows.append(
                    ft.DataRow(
                        cells=[
                            ft.DataCell(ft.Text(str(row['Entity']) if pd.notna(row['Entity']) else "-", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['ProjectNo']) if pd.notna(row['ProjectNo']) else "-", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['WONo']), color=COLORS['accent'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['SO_FCST']) if pd.notna(row['SO_FCST']) else "-", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['Sub']) if pd.notna(row['Sub']) else "-", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['ParentWO']) if pd.notna(row['ParentWO']) else "-", color=COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['ItemNumber']) if pd.notna(row['ItemNumber']) else "-", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['Description']) if pd.notna(row['Description']) else "-", color=COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['DueDt'])[:10] if pd.notna(row['DueDt']) else "-", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['WoType']) if pd.notna(row['WoType']) else "-", color=COLORS['text_secondary'], selectable=True)),
                            ft.DataCell(ft.Text(wo_status, color=status_color, selectable=True)),
                            # COLUMNAS DE OPERACIÓN CON NOMBRES CORRECTOS
                            ft.DataCell(ft.Text(str(row['ItemNo']) if pd.notna(row['ItemNo']) else "-", color=COLORS['warning'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['Seq']) if pd.notna(row['Seq']) else "-", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(str(row['OpID']) if pd.notna(row['OpID']) else "-", color=COLORS['text_primary'], selectable=True)),
                            ft.DataCell(ft.Text(op_status, color=op_status_color, selectable=True)),
                            ft.DataCell(ft.Text(str(row['WrkCtr']) if pd.notna(row['WrkCtr']) else "-", color=COLORS['text_secondary'], selectable=True)),
                        ]
                    )
                )
            
            # Actualizar status con información de operaciones
            total_wos = len(df_wo_inquiry['WONo'].unique()) if 'WONo' in df_wo_inquiry.columns else 0
            total_operations = len(df_wo_inquiry[df_wo_inquiry['Seq'].notna()]) if 'Seq' in df_wo_inquiry.columns else 0
            related_wos = len(df_wo_inquiry[df_wo_inquiry['ParentWO'] == searched_wo]) if 'ParentWO' in df_wo_inquiry.columns else 0
            direct_wos = len(df_wo_inquiry[df_wo_inquiry['WONo'] == searched_wo]) if 'WONo' in df_wo_inquiry.columns else 0
            
            wo_inquiry_status.current.value = f"✅ Encontradas {total_wos} WOs únicas | {direct_wos} directas | {related_wos} relacionadas | {total_operations} operaciones"
            wo_inquiry_status.current.color = COLORS['success']
            wo_inquiry_container.current.visible = True
        
        wo_inquiry_table.current.update()
        wo_inquiry_container.current.update()
        wo_inquiry_status.current.update()

    def search_wo_materials(e):
        """Buscar todos los materiales de una WO desde tabla PR561 y WOs relacionadas desde WOInquiry"""
        wo_number = wo_search_entry.current.value.strip() if wo_search_entry.current.value else ""
        
        if not wo_number:
            wo_search_status.current.value = "❌ Ingresa un número de WO"
            wo_search_status.current.color = COLORS['error']
            wo_search_status.current.update()
            return
        
        wo_search_button.current.text = "🔍 Buscando..."
        wo_search_button.current.disabled = True
        wo_search_status.current.value = "🔄 Consultando base de datos..."
        wo_search_status.current.color = COLORS['text_secondary']
        wo_search_button.current.update()
        wo_search_status.current.update()
        
        try:
            conn = sqlite3.connect(dashboard.db_path)
            
            # Consultar materiales de la WO desde PR561 con costos de IN92
            query = """
            SELECT pr.ItemNo as Component, pr.ReqDate, 
                CAST(COALESCE(pr.QtyOh, 0) AS REAL) as QtyOh,
                CAST(COALESCE(pr.ReqQty, 0) AS REAL) as ReqQty,
                CAST(COALESCE(pr.QtyPending, 0) AS REAL) as QtyPending,
                pr.Srt, pr.Project, pr.Entity, pr.Description,
                CAST(COALESCE(in92.STDCost, 0) AS REAL) as STDCost
            FROM pr561 pr 
            LEFT JOIN in92 ON pr.ItemNo = in92.ItemNo 
            WHERE pr.WONo = ?
            ORDER BY pr.ReqDate, pr.ItemNo
            """
            
            df_materials = pd.read_sql_query(query, conn, params=[wo_number])
            conn.close()
            
            if len(df_materials) == 0:
                wo_search_status.current.value = f"⚠️ No se encontraron materiales para WO: {wo_number}"
                wo_search_status.current.color = COLORS['warning']
                wo_materials_container.current.visible = False
                wo_inquiry_container.current.visible = False
                wo_materials_container.current.update()
                wo_inquiry_container.current.update()
            else:
                # Limpiar tabla anterior
                wo_materials_table.current.rows.clear()
                
                # Asegurar tipos numéricos
                df_materials['QtyOh'] = pd.to_numeric(df_materials['QtyOh'], errors='coerce').fillna(0)
                df_materials['ReqQty'] = pd.to_numeric(df_materials['ReqQty'], errors='coerce').fillna(0)
                df_materials['QtyPending'] = pd.to_numeric(df_materials['QtyPending'], errors='coerce').fillna(0)
                df_materials['STDCost'] = pd.to_numeric(df_materials['STDCost'], errors='coerce').fillna(0)
                
                # Calcular valores
                df_materials['Valor_Pendiente'] = df_materials['QtyPending'] * df_materials['STDCost']
                df_materials['Valor_Total'] = df_materials['ReqQty'] * df_materials['STDCost']
                
                # Llenar tabla con materiales encontrados
                for _, row in df_materials.iterrows():
                    # Determinar color del status
                    qty_pending = float(row['QtyPending'])
                    if qty_pending == 0:
                        status_color = COLORS['success']
                        status_text = "Surtido"
                    else:
                        status_color = COLORS['warning']
                        status_text = "Pendiente"
                    
                    wo_materials_table.current.rows.append(
                        ft.DataRow(
                            cells=[
                                ft.DataCell(ft.Text(str(row['Component']), color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(str(row['ReqDate'])[:10], color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(f"{float(row['QtyOh']):,.0f}", color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(f"{float(row['ReqQty']):,.0f}", color=COLORS['accent'], selectable=True)),
                                ft.DataCell(ft.Text(f"{qty_pending:,.0f}", color=status_color, selectable=True)),
                                ft.DataCell(ft.Text(status_text, color=status_color, selectable=True)),
                                ft.DataCell(ft.Text(f"${float(row['STDCost']):,.2f}", color=COLORS['text_primary'], selectable=True)),
                                ft.DataCell(ft.Text(f"${float(row['Valor_Pendiente']):,.2f}", color=COLORS['error'], selectable=True)),
                                ft.DataCell(ft.Text(f"${float(row['Valor_Total']):,.2f}", color=COLORS['success'], selectable=True)),
                                ft.DataCell(ft.Text(str(row['Srt']) if pd.notna(row['Srt']) and str(row['Srt']) != 'nan' else "-", color=COLORS['text_secondary'], selectable=True)),
                            ]
                        )
                    )
                
                # Calcular resumen de materiales
                total_components = len(df_materials)
                pending_components = len(df_materials[df_materials['QtyPending'] > 0])
                total_value_pending = float(df_materials['Valor_Pendiente'].sum())
                total_value_all = float(df_materials['Valor_Total'].sum())
                
                wo_search_status.current.value = f"✅ Encontrados {total_components} componentes | {pending_components} pendientes | Valor pendiente: ${total_value_pending:,.2f}"
                wo_search_status.current.color = COLORS['success']
                wo_materials_container.current.visible = True
                
                # NUEVO: Buscar WOs relacionadas en WOInquiry
                df_wo_inquiry = search_wo_inquiry_data(wo_number)
                update_wo_inquiry_table(df_wo_inquiry, wo_number)
                
                wo_materials_table.current.update()
                wo_materials_container.current.update()
            
        except Exception as e:
            wo_search_status.current.value = f"❌ Error: {str(e)}"
            wo_search_status.current.color = COLORS['error']
            wo_materials_container.current.visible = False
            wo_inquiry_container.current.visible = False
            wo_materials_container.current.update()
            wo_inquiry_container.current.update()
            print(f"Error detallado en search_wo_materials: {e}")
        
        wo_search_button.current.text = "🔍 Buscar Materiales"
        wo_search_button.current.disabled = False
        wo_search_button.current.update()
        wo_search_status.current.update()
    
    def download_wo_clear_checked(e):
        """Descargar tabla WOClear_checked a Excel"""
        download_button.current.text = "⏳ Descargando..."
        download_button.current.disabled = True
        download_button.current.update()
        
        try:
            conn = sqlite3.connect(dashboard.db_path)
            
            # Consultar todos los registros de WOClear_checked
            query = """
            SELECT WONo, checked_date, user_action, notes, 
                datetime(checked_date, 'localtime') as fecha_registro
            FROM WOClear_checked 
            ORDER BY checked_date DESC
            """
            
            df_wo_checked = pd.read_sql_query(query, conn)
            conn.close()
            
            if len(df_wo_checked) == 0:
                register_status.current.value = "⚠️ No hay registros en WOClear_checked para descargar"
                register_status.current.color = COLORS['warning']
            else:
                # Crear nombre de archivo con timestamp
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"WOClear_Checked_{timestamp}.xlsx"
                
                # Renombrar columnas para mejor presentación
                df_wo_checked.columns = [
                    'Número WO', 'Fecha Registro (UTC)', 'Acción Usuario', 
                    'Notas', 'Fecha Registro (Local)'
                ]
                
                # Exportar a Excel
                df_wo_checked.to_excel(filename, index=False, engine='openpyxl')
                
                register_status.current.value = f"✅ Descargado: {filename} ({len(df_wo_checked)} registros)"
                register_status.current.color = COLORS['success']
                
        except Exception as e:
            register_status.current.value = f"❌ Error descargando: {str(e)}"
            register_status.current.color = COLORS['error']
            print(f"Error detallado en descarga: {e}")
        
        download_button.current.text = "📥 Descargar Excel"
        download_button.current.disabled = False
        download_button.current.update()
        register_status.current.update()
    
    def register_wo_checked(e):
        """Registrar WO en tabla WOClear_checked con notas"""
        wo_number = wo_entry.current.value.strip() if wo_entry.current.value else ""
        notes = wo_notes.current.value.strip() if wo_notes.current.value else ""
        
        if not wo_number:
            register_status.current.value = "❌ Ingresa un número de WO"
            register_status.current.color = COLORS['error']
            register_status.current.update()
            return
        
        register_button.current.text = "⏳ Registrando..."
        register_button.current.disabled = True
        register_button.current.update()
        
        try:
            # Crear tabla si no existe
            create_wo_clear_checked_table()
            
            conn = sqlite3.connect(dashboard.db_path)
            
            # Verificar si la WO existe en WOClear
            check_wo_query = "SELECT COUNT(*) FROM WOClear WHERE WONo = ?"
            wo_exists = conn.execute(check_wo_query, (wo_number,)).fetchone()[0] > 0
            
            if not wo_exists:
                register_status.current.value = f"⚠️ WO {wo_number} no existe en WOClear"
                register_status.current.color = COLORS['warning']
            else:
                # Insertar o actualizar registro con notas
                insert_query = """
                INSERT OR REPLACE INTO WOClear_checked (WONo, checked_date, user_action, notes)
                VALUES (?, CURRENT_TIMESTAMP, 'REGISTERED', ?)
                """
                conn.execute(insert_query, (wo_number, notes))
                conn.commit()
                
                register_status.current.value = f"✅ WO {wo_number} registrada exitosamente"
                register_status.current.color = COLORS['success']
                wo_entry.current.value = ""  # Limpiar campo WO
                wo_notes.current.value = ""  # Limpiar campo notas
                wo_entry.current.update()
                wo_notes.current.update()
            
            conn.close()
            
        except sqlite3.IntegrityError:
            register_status.current.value = f"⚠️ WO {wo_number} ya está registrada"
            register_status.current.color = COLORS['warning']
        except Exception as e:
            register_status.current.value = f"❌ Error: {str(e)}"
            register_status.current.color = COLORS['error']
        
        register_button.current.text = "📝 Registrar WO"
        register_button.current.disabled = False
        register_button.current.update()
        register_status.current.update()

    def get_top_80_wos_data(pareto_data):
        """Obtener el 80% de las WOs más costosas con información adicional de WOInquiry"""
        if pareto_data is None or len(pareto_data) == 0:
            return pd.DataFrame()
        
        # Calcular cuántas WOs representan el 80%
        cumulative_80_idx = pareto_data[pareto_data['Pct_Acumulado'] <= 80].index
        if len(cumulative_80_idx) == 0:
            # Si ninguna llega al 80%, tomar al menos el top 20%
            top_80_count = max(1, int(len(pareto_data) * 0.2))
        else:
            top_80_count = cumulative_80_idx[-1] + 1
        
        # Obtener datos del 80%
        top_80_data = pareto_data.head(top_80_count).copy()
        
        # NUEVO: Enriquecer con datos de WOInquiry
        try:
            conn = sqlite3.connect(dashboard.db_path)
            
            # Obtener información adicional de WOInquiry para cada WO
            wo_info = {}
            for _, row in top_80_data.iterrows():
                wo_number = str(row['WONo'])
                
                query = """
                SELECT ItemNumber, Description 
                FROM WOInquiry 
                WHERE WONo = ? 
                LIMIT 1
                """
                
                result = pd.read_sql_query(query, conn, params=[wo_number])
                
                if len(result) > 0:
                    wo_info[wo_number] = {
                        'ItemNumber': str(result.iloc[0]['ItemNumber']) if pd.notna(result.iloc[0]['ItemNumber']) else "-",
                        'Description': str(result.iloc[0]['Description']) if pd.notna(result.iloc[0]['Description']) else "-"
                    }
                else:
                    wo_info[wo_number] = {'ItemNumber': "-", 'Description': "-"}
            
            conn.close()
            
            # Agregar columnas de Item y Description
            top_80_data['ItemNumber'] = top_80_data['WONo'].astype(str).map(lambda x: wo_info.get(x, {}).get('ItemNumber', '-'))
            top_80_data['Description'] = top_80_data['WONo'].astype(str).map(lambda x: wo_info.get(x, {}).get('Description', '-'))
            
        except Exception as e:
            print(f"Error enriqueciendo datos con WOInquiry: {e}")
            # Si hay error, agregar columnas vacías
            top_80_data['ItemNumber'] = "-"
            top_80_data['Description'] = "-"
        
        return top_80_data
    
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
                # LIMPIAR TODAS LAS TABLAS cuando no hay datos
                top_80_table.current.rows.clear()
                wo_materials_table.current.rows.clear()
                wo_inquiry_table.current.rows.clear()
                wo_materials_container.current.visible = False
                wo_inquiry_container.current.visible = False
                
                # Ocultar gráfico si no hay datos
                pareto_container.current.visible = False
                
                wo_clear_metrics.current.update()
                top_80_table.current.update()
                wo_materials_table.current.update()
                wo_inquiry_table.current.update()
                wo_materials_container.current.update()
                wo_inquiry_container.current.update()
                pareto_container.current.update()
                return
            
            # Aplicar filtros
            df_filtered = df_wo_clear.copy()
            
            if selected_wo.current and selected_wo.current.value and selected_wo.current.value != "Todas":
                df_filtered = df_filtered[df_filtered['WONo'] == selected_wo.current.value]
            
            if selected_project.current and selected_project.current.value and selected_project.current.value != "Todos":
                df_filtered = df_filtered[df_filtered['Project'] == selected_project.current.value]
            
            if selected_entity.current and selected_entity.current.value and selected_entity.current.value != "Todas":
                df_filtered = df_filtered[df_filtered['Entity'] == selected_entity.current.value]
            
            # Actualizar métricas
            total_wos = len(df_filtered['WONo'].unique()) if len(df_filtered) > 0 else 0
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
🎯 WOs críticas: {', '.join(pareto_data.head(3)['WONo'].astype(str))}"""
                    
                    pareto_container.current.content.controls[1] = ft.Text(
                        analysis_text.strip(),
                        size=12, 
                        color=COLORS['text_secondary'],
                        text_align=ft.TextAlign.CENTER
                    )
                    
                    # Actualizar tabla del 80%
                    update_top_80_table(pareto_data)
                else:
                    pareto_container.current.visible = False
            else:
                pareto_container.current.visible = False
            
            # Limpiar tablas de materiales y WO inquiry
            wo_materials_table.current.rows.clear()
            wo_inquiry_table.current.rows.clear()
            wo_materials_container.current.visible = False
            wo_inquiry_container.current.visible = False
            
            wo_clear_metrics.current.update()
            top_80_table.current.update()
            wo_materials_table.current.update()
            wo_inquiry_table.current.update()
            wo_materials_container.current.update()
            wo_inquiry_container.current.update()
            pareto_container.current.update()
            pareto_image.current.update()
            
        except Exception as e:
            print(f"Error cargando WO Clear: {e}")
    
    def update_top_80_table(pareto_data):
        """Actualizar tabla del 80% de WOs más costosas CON COLUMNAS REORDENADAS"""
        top_80_data = get_top_80_wos_data(pareto_data)
        
        # Limpiar tabla
        top_80_table.current.rows.clear()
        
        for idx, row in top_80_data.iterrows():
            # Determinar color según posición en el ranking
            if row['Ranking'] <= 3:
                ranking_color = COLORS['error']  # Top 3 en rojo
            elif row['Ranking'] <= 10:
                ranking_color = COLORS['warning']  # Top 10 en amarillo
            else:
                ranking_color = COLORS['success']  # Resto en verde
            
            top_80_table.current.rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(f"#{row['Ranking']}", color=ranking_color, weight=ft.FontWeight.BOLD, selectable=True)),
                        ft.DataCell(ft.Text(str(row['WONo']), color=COLORS['text_primary'], selectable=True)),
                        ft.DataCell(ft.Text(str(row['ItemNumber']), color=COLORS['accent'], selectable=True)),  # MOVIDA AQUÍ
                        ft.DataCell(ft.Text(f"${row['Total_Value']:,.0f}", color=COLORS['success'], selectable=True)),
                        ft.DataCell(ft.Text(f"{row['Pct_Individual']:.1f}%", color=COLORS['accent'], selectable=True)),
                        ft.DataCell(ft.Text(f"{row['Pct_Acumulado']:.1f}%", color=COLORS['text_secondary'], selectable=True)),
                        ft.DataCell(ft.Text(f"{row['Component_Count']}", color=COLORS['text_primary'], selectable=True)),
                        ft.DataCell(ft.Text(str(row['Project']) if row['Project'] else "-", color=COLORS['text_secondary'], selectable=True)),
                        ft.DataCell(ft.Text(str(row['Description'])[:25] + "..." if len(str(row['Description'])) > 25 else str(row['Description']), 
                                        color=COLORS['text_secondary'], selectable=True)),  # AL FINAL Y MÁS CORTA
                    ]
                )
            )
    
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
    
    # Sección de registro de WOs CON CAMPO DE NOTAS Y BOTÓN DE DESCARGA
    wo_register_section = ft.Container(
        content=ft.Column([
            ft.Text("📝 Registrar WO para Acción", 
                size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
            ft.Text("Registra WOs que requieren seguimiento", 
                size=12, color=COLORS['text_secondary']),
            ft.Row([
                ft.TextField(
                    ref=wo_entry,
                    label="Número de WO",
                    hint_text="Ej: WO-2024-001",
                    width=200,
                    bgcolor=COLORS['secondary'],
                    color=COLORS['text_primary']
                ),
                ft.ElevatedButton(
                    "📝 Registrar WO",
                    ref=register_button,
                    on_click=register_wo_checked,
                    bgcolor=COLORS['success'],
                    color=COLORS['text_primary'],
                    icon=ft.Icons.CHECK_CIRCLE
                ),
                ft.ElevatedButton(
                    "📥 Descargar Excel",
                    ref=download_button,
                    on_click=download_wo_clear_checked,
                    bgcolor=COLORS['warning'],
                    color=COLORS['text_primary'],
                    icon=ft.Icons.DOWNLOAD
                )
            ], spacing=10),
            # Campo de notas
            ft.TextField(
                ref=wo_notes,
                label="Notas (Opcional)",
                hint_text="Agregar comentarios o acciones a realizar...",
                multiline=True,
                min_lines=2,
                max_lines=4,
                bgcolor=COLORS['secondary'],
                color=COLORS['text_primary']
            ),
            ft.Text("", ref=register_status, size=12)
        ], spacing=10),
        padding=ft.padding.all(15),
        bgcolor=COLORS['primary'],
        border_radius=8,
        border=ft.border.all(1, COLORS['accent'])
    )
    
    # Tabla del 80% de WOs más costosas
    top_80_wo_table = ft.DataTable(
        ref=top_80_table,
        columns=[
            ft.DataColumn(ft.Text("Rank", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("WO Number", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Item Number", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Total Value", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("% Individual", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("% Acumulado", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Components", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Project", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Description", color=COLORS['text_primary'])),
        ],
        rows=[],
        bgcolor=COLORS['primary'],
        border_radius=8,
        column_spacing=10,
    )
    
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
    
    # Contenedor de métricas
    metrics_container = ft.Column(ref=wo_clear_metrics, controls=[])
    
    # Tabla de materiales de WO específica
    wo_materials_data_table = ft.DataTable(
        ref=wo_materials_table,
        columns=[
            ft.DataColumn(ft.Text("Component", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Req Date", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Qty OH", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Req Qty", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Qty Pending", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Status", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("STD Cost", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Valor Pendiente", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Valor Total", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Sort Code", color=COLORS['text_primary'])),
        ],
        rows=[],
        bgcolor=COLORS['primary'],
        border_radius=8,
    )
    
    # Tabla de WOInquiry con operaciones
    wo_inquiry_data_table = ft.DataTable(
        ref=wo_inquiry_table,
        columns=[
            ft.DataColumn(ft.Text("Entity", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Project", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("WO Number", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("SO/FCST", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Sub", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Parent WO", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Item Number", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Description", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Due Date", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("WO Type", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("WO Status", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Op Item", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Seq", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Op ID", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Op Status", color=COLORS['text_primary'])),
            ft.DataColumn(ft.Text("Work Center", color=COLORS['text_primary'])),
        ],
        rows=[],
        bgcolor=COLORS['primary'],
        border_radius=8,
        column_spacing=8,
    )
    
    # Contenedor para tabla de materiales de WO (inicialmente oculto)
    wo_materials_detail_container = ft.Container(
        ref=wo_materials_container,
        visible=False,
        content=ft.Column([
            ft.Text("📋 Materiales de Work Order", 
                size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
            ft.Text("Todos los valores son seleccionables para copiar", 
                size=12, color=COLORS['text_secondary']),
            ft.Container(
                content=ft.Column([
                    wo_materials_data_table
                ], scroll=ft.ScrollMode.ALWAYS),
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
    
    # Contenedor para tabla de WOInquiry (inicialmente oculto)
    wo_inquiry_detail_container = ft.Container(
        ref=wo_inquiry_container,
        visible=False,
        content=ft.Column([
            ft.Text("🔗 Work Orders Relacionadas con Operaciones", 
                size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
            ft.Text("", ref=wo_inquiry_status, size=12),
            ft.Text("WOs donde aparece como WONo o ParentWO + sus operaciones de Operation_WO (Seleccionable)", 
                size=12, color=COLORS['text_secondary']),
            ft.Container(
                content=ft.Column([
                    wo_inquiry_data_table
                ], scroll=ft.ScrollMode.ALWAYS),
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
        
        # Contenedor de métricas
        ft.Container(
            content=metrics_container,
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        ),
        
        # Gráfico de Pareto
        pareto_chart_container,
        
        # Tabla del 80% con búsqueda y registro
        ft.Container(
            content=ft.Column([
                ft.Text("🏆 Top 80% WOs Más Costosas", 
                    size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Text("WOs que concentran el mayor valor (Scrollable y Seleccionable)", 
                    size=12, color=COLORS['text_secondary']),
                ft.Container(
                    content=ft.Column([
                        top_80_wo_table
                    ], scroll=ft.ScrollMode.ALWAYS),
                    border=ft.border.all(1, COLORS['secondary']),
                    border_radius=8,
                    padding=15,
                    height=500,
                    bgcolor=COLORS['primary'],
                ),
                
                # Sección de búsqueda y registro
                ft.Row([
                    # Búsqueda de materiales
                    ft.Container(
                        content=ft.Column([
                            ft.Text("🔍 Buscar Materiales de WO", 
                                    size=14, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                            ft.Text("Consulta todos los materiales de una Work Order desde PR561", 
                                    size=11, color=COLORS['text_secondary']),
                            ft.Row([
                                ft.TextField(
                                    ref=wo_search_entry,
                                    label="Número de WO",
                                    hint_text="Ej: 12073358",
                                    width=200,
                                    bgcolor=COLORS['secondary'],
                                    color=COLORS['text_primary']
                                ),
                                ft.ElevatedButton(
                                    "🔍 Buscar Materiales",
                                    ref=wo_search_button,
                                    on_click=search_wo_materials,
                                    bgcolor=COLORS['accent'],
                                    color=COLORS['text_primary'],
                                    icon=ft.Icons.SEARCH
                                )
                            ], spacing=10),
                            ft.Text("", ref=wo_search_status, size=11)
                        ], spacing=8),
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['secondary'],
                        border_radius=8,
                        border=ft.border.all(1, COLORS['accent']),
                        expand=1
                    ),
                    
                    # Sección de registro
                    wo_register_section
                    
                ], spacing=20)
                
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        ),
        
        # Contenedor de materiales de WO específica
        wo_materials_detail_container,
        
        # Contenedor de WOInquiry
        wo_inquiry_detail_container,
        
    ], spacing=20, scroll=ft.ScrollMode.AUTO)

def main(page: ft.Page):
    page.title = "Cobertura de Inventario WO Clear"
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
                ft.DataColumn(ft.Text("Item No", color=COLORS['text_primary'])),
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
                                ft.Text("📊 Cobertura de Inventario WO Clear", 
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
                                    ft.DataCell(ft.Text(str(row['ItemNo']), color=COLORS['text_primary'])),
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