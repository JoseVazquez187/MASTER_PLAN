import flet as ft
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import io
import base64
import sqlite3
import math
import os
import subprocess
import platform
import matplotlib
import tempfile
import asyncio

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

class ObsoleteSlowMotionAnalyzer:
    def __init__(self):
        self.slowmotion_data = None
        self.obsolete_items_data = None
        self.new_items = None
        self.metrics = {}
        
    def load_data_from_db(self, db_path=r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"):
        """Cargar datos desde R4Database"""
        try:
            print(f"Conectando a R4Database: {db_path}")
            conn = sqlite3.connect(db_path)
            
            # Verificar tablas disponibles
            tables_query = "SELECT name FROM sqlite_master WHERE type='table'"
            tables = pd.read_sql_query(tables_query, conn)
            print(f"📂 Tablas disponibles: {tables['name'].tolist()}")
            
            # Cargar slowmotion (items obsoletos)
            print("🔍 Cargando tabla slowmotion...")
            slowmotion_query = """
            SELECT id, ItemNo, Description, UM, S, WH, LastDate, Bin, Lot, Rev, Entity,
                   PlanType, Prj, PrdGrp, MLI, LT, PlnrCd, ABCD, BM, Cost, QtyOH, ExtOH
            FROM slowmotion
            """
            self.slowmotion_data = pd.read_sql_query(slowmotion_query, conn)
            print(f"📦 Slowmotion items cargados: {len(self.slowmotion_data)}")
            
            # Verificar si existe la tabla newObsolet_slowmotion_items
            table_exists_query = """
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='newObsolet_slowmotion_items'
            """
            table_exists = pd.read_sql_query(table_exists_query, conn)
            
            if len(table_exists) == 0:
                print("🆕 Creando tabla newObsolet_slowmotion_items...")
                self._create_obsolete_items_table(conn)
                # Insertar datos iniciales de slowmotion
                self._initialize_obsolete_items_table(conn)
            else:
                print("📋 Cargando tabla newObsolet_slowmotion_items existente...")
                obsolete_query = """
                SELECT * FROM newObsolet_slowmotion_items
                """
                self.obsolete_items_data = pd.read_sql_query(obsolete_query, conn)
                print(f"📊 Items obsoletos registrados: {len(self.obsolete_items_data)}")
            
            conn.close()
            
            # Procesar datos y calcular métricas
            self._process_data()
            return True
            
        except Exception as e:
            print(f"❌ Error cargando desde DB: {e}")
            print("🔄 Usando datos de ejemplo...")
            self._load_sample_data()
            return False
    
    def _create_obsolete_items_table(self, conn):
        """Crear tabla newObsolet_slowmotion_items con campos adicionales"""
        create_table_query = """
        CREATE TABLE IF NOT EXISTS newObsolet_slowmotion_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ItemNo TEXT NOT NULL,
            Description TEXT,
            UM TEXT,
            S TEXT,
            WH TEXT,
            LastDate TEXT,
            Bin TEXT,
            Lot TEXT,
            Rev TEXT,
            Entity TEXT,
            PlanType TEXT,
            Prj TEXT,
            PrdGrp TEXT,
            MLI TEXT,
            LT INTEGER,
            PlnrCd TEXT,
            ABCD TEXT,
            BM TEXT,
            Cost REAL,
            QtyOH REAL,
            ExtOH REAL,
            date_added TEXT DEFAULT CURRENT_TIMESTAMP,
            root_cause TEXT DEFAULT '',
            status TEXT DEFAULT 'Pending Analysis',
            engineer_notes TEXT DEFAULT '',
            is_new INTEGER DEFAULT 1,
            department TEXT DEFAULT '',
            responsible_person TEXT DEFAULT '',
            assigned_date TEXT DEFAULT '',
            UNIQUE(ItemNo, Lot, Bin)
        )
        """
        conn.execute(create_table_query)
        conn.commit()
        print("✅ Tabla newObsolet_slowmotion_items creada con campos adicionales")
        

    def _initialize_obsolete_items_table(self, conn):
        """Inicializar tabla con datos de slowmotion - SIN DUPLICADOS"""
        if self.slowmotion_data is not None and len(self.slowmotion_data) > 0:
            
            # VERIFICAR SI YA HAY DATOS
            check_query = "SELECT COUNT(*) as count FROM newObsolet_slowmotion_items"
            result = pd.read_sql_query(check_query, conn)
            existing_count = result['count'].iloc[0]
            
            if existing_count > 0:
                print(f"✅ Tabla ya tiene {existing_count} registros, saltando inicialización")
                # Cargar los datos existentes
                obsolete_query = "SELECT * FROM newObsolet_slowmotion_items"
                self.obsolete_items_data = pd.read_sql_query(obsolete_query, conn)
                return
            
            # Si no hay datos, proceder con la inicialización
            init_data = self.slowmotion_data.copy()
            init_data['date_added'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            init_data['root_cause'] = ''
            init_data['status'] = 'Historical'
            init_data['engineer_notes'] = ''
            init_data['is_new'] = 0
            
            # Insertar datos
            init_data.to_sql('newObsolet_slowmotion_items', conn, if_exists='append', index=False)
            
            # Cargar los datos insertados
            obsolete_query = "SELECT * FROM newObsolet_slowmotion_items"
            self.obsolete_items_data = pd.read_sql_query(obsolete_query, conn)
            
            print(f"✅ Inicializada tabla con {len(init_data)} items históricos")

    def _load_sample_data(self):
        """Crear datos de ejemplo para testing"""
        np.random.seed(42)
        n_items = 150
        
        # Datos simulados de slowmotion
        slowmotion_data = {
            'id': range(1, n_items + 1),
            'ItemNo': [f"SM-{1000+i:04d}" for i in range(n_items)],
            'Description': [f"Slow Motion Item {i}" for i in range(n_items)],
            'UM': np.random.choice(['EA', 'LB', 'FT', 'KG'], n_items),
            'S': np.random.choice(['A', 'I', 'O'], n_items),
            'WH': np.random.choice(['WH01', 'WH02', 'WH03'], n_items),
            'LastDate': pd.date_range(start='2023-01-01', periods=n_items, freq='D'),
            'Bin': [f"BIN-{i:03d}" for i in range(n_items)],
            'Lot': [f"LOT-{i:04d}" for i in range(n_items)],
            'Rev': np.random.choice(['A', 'B', 'C', ''], n_items),
            'Entity': np.random.choice(['MX', 'US', 'CA'], n_items),
            'PlanType': np.random.choice(['MTO', 'MTS', 'BUY', 'MAKE'], n_items),
            'Prj': np.random.choice(['PROJ01', 'PROJ02', 'PROJ03'], n_items),
            'PrdGrp': np.random.choice(['GRP01', 'GRP02', 'GRP03'], n_items),
            'MLI': np.random.choice(['L', 'M', 'H'], n_items),
            'LT': np.random.randint(1, 90, n_items),
            'PlnrCd': np.random.choice(['PLN01', 'PLN02', 'PLN03'], n_items),
            'ABCD': np.random.choice(['A', 'B', 'C', 'D'], n_items),
            'BM': np.random.choice(['BM1', 'BM2', 'BM3'], n_items),
            'Cost': np.random.uniform(10, 1000, n_items),
            'QtyOH': np.random.randint(1, 500, n_items),
            'ExtOH': np.random.uniform(100, 50000, n_items)
        }
        self.slowmotion_data = pd.DataFrame(slowmotion_data)
        
        # Crear datos de obsolete items (simulando tabla existente)
        obsolete_data = self.slowmotion_data.copy()
        obsolete_data['date_added'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        obsolete_data['root_cause'] = ''
        obsolete_data['status'] = 'Historical'
        obsolete_data['engineer_notes'] = ''
        obsolete_data['is_new'] = 0
        
        # Agregar algunos items "nuevos" para simular
        new_items_count = 10
        new_items_data = {
            'id': range(n_items + 1, n_items + new_items_count + 1),
            'ItemNo': [f"NEW-{2000+i:04d}" for i in range(new_items_count)],
            'Description': [f"New Obsolete Item {i}" for i in range(new_items_count)],
            'UM': np.random.choice(['EA', 'LB', 'FT'], new_items_count),
            'S': np.random.choice(['A', 'I', 'O'], new_items_count),
            'WH': np.random.choice(['WH01', 'WH02'], new_items_count),
            'LastDate': pd.date_range(start='2024-12-01', periods=new_items_count, freq='D'),
            'Bin': [f"NBIN-{i:03d}" for i in range(new_items_count)],
            'Lot': [f"NLOT-{i:04d}" for i in range(new_items_count)],
            'Rev': np.random.choice(['A', 'B'], new_items_count),
            'Entity': np.random.choice(['MX', 'US'], new_items_count),
            'PlanType': np.random.choice(['MTO', 'MTS'], new_items_count),
            'Prj': np.random.choice(['PROJ01', 'PROJ02'], new_items_count),
            'PrdGrp': np.random.choice(['GRP01', 'GRP02'], new_items_count),
            'MLI': np.random.choice(['L', 'M'], new_items_count),
            'LT': np.random.randint(1, 60, new_items_count),
            'PlnrCd': np.random.choice(['PLN01', 'PLN02'], new_items_count),
            'ABCD': np.random.choice(['A', 'B', 'C'], new_items_count),
            'BM': np.random.choice(['BM1', 'BM2'], new_items_count),
            'Cost': np.random.uniform(10, 500, new_items_count),
            'QtyOH': np.random.randint(1, 100, new_items_count),
            'ExtOH': np.random.uniform(100, 10000, new_items_count),
            'date_added': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'root_cause': '',
            'status': 'Pending Analysis',
            'engineer_notes': '',
            'is_new': 1
        }
        
        new_items_df = pd.DataFrame(new_items_data)
        self.obsolete_items_data = pd.concat([obsolete_data, new_items_df], ignore_index=True)
        
        print("🧪 Datos de ejemplo creados")
        self._process_data()
    
    def _process_data(self):
        """Procesar datos y calcular métricas"""
        try:
            if self.slowmotion_data is None or len(self.slowmotion_data) == 0:
                print("⚠️ No hay datos de slowmotion para procesar")
                return
            
            # Limpiar datos
            self.slowmotion_data['PlanType'] = self.slowmotion_data['PlanType'].fillna('Unknown')
            self.slowmotion_data['ExtOH'] = pd.to_numeric(self.slowmotion_data['ExtOH'], errors='coerce').fillna(0)
            
            if self.obsolete_items_data is not None:
                self.obsolete_items_data['PlanType'] = self.obsolete_items_data['PlanType'].fillna('Unknown')
                self.obsolete_items_data['ExtOH'] = pd.to_numeric(self.obsolete_items_data['ExtOH'], errors='coerce').fillna(0)
                
                # Identificar items nuevos
                self.new_items = self.obsolete_items_data[
                    self.obsolete_items_data['is_new'] == 1
                ].copy()
            
            self._calculate_metrics()
            print("✅ Procesamiento completado exitosamente")
            
        except Exception as e:
            print(f"❌ Error procesando datos: {e}")
            import traceback
            traceback.print_exc()
    
    def _calculate_metrics(self):
        """Calcular métricas principales"""
        if self.slowmotion_data is None or len(self.slowmotion_data) == 0:
            self.metrics = {}
            return
        
        # Métricas de slowmotion
        total_items = len(self.slowmotion_data)
        total_value = self.slowmotion_data['ExtOH'].sum()
        
        # Análisis por PlanType
        plantype_analysis = self.slowmotion_data.groupby('PlanType').agg({
            'ExtOH': 'sum',
            'ItemNo': 'count'
        }).reset_index()
        plantype_analysis.columns = ['PlanType', 'Total_Value', 'Item_Count']
        plantype_analysis = plantype_analysis.sort_values('Total_Value', ascending=False)
        
        # Métricas de items obsoletos
        new_items_count = 0
        pending_analysis_count = 0
        avg_days_in_system = 0
        
        if self.obsolete_items_data is not None:
            new_items_count = len(self.obsolete_items_data[self.obsolete_items_data['is_new'] == 1])
            pending_analysis_count = len(self.obsolete_items_data[
                self.obsolete_items_data['status'] == 'Pending Analysis'
            ])
            
            # Calcular días promedio en el sistema
            if 'date_added' in self.obsolete_items_data.columns:
                self.obsolete_items_data['date_added'] = pd.to_datetime(
                    self.obsolete_items_data['date_added'], errors='coerce'
                )
                current_date = datetime.now()
                days_diff = (current_date - self.obsolete_items_data['date_added']).dt.days
                avg_days_in_system = days_diff.mean()
        
        # Top entities y projects
        top_entity = 'N/A'
        top_project = 'N/A'
        
        if 'Entity' in self.slowmotion_data.columns:
            entity_mode = self.slowmotion_data['Entity'].mode()
            if len(entity_mode) > 0:
                top_entity = entity_mode.iloc[0]
                
        if 'Prj' in self.slowmotion_data.columns:
            project_mode = self.slowmotion_data['Prj'].mode()
            if len(project_mode) > 0:
                top_project = project_mode.iloc[0]
        
        self.metrics = {
            'total_items': total_items,
            'total_value': round(total_value, 2),
            'plantype_analysis': plantype_analysis,
            'new_items_count': new_items_count,
            'pending_analysis_count': pending_analysis_count,
            'avg_days_in_system': round(avg_days_in_system, 1) if not np.isnan(avg_days_in_system) else 0,
            'top_entity': top_entity,
            'top_project': top_project,
            'unique_entities': self.slowmotion_data['Entity'].nunique(),
            'unique_projects': self.slowmotion_data['Prj'].nunique()
        }
    
    def create_plantype_chart(self):
        """Crear Diagrama de Pareto por PlanType (80% de los valores)"""
        if self.slowmotion_data is None or len(self.slowmotion_data) == 0:
            return None
        
        try:
            import matplotlib
            matplotlib.use('Agg')
            
            plt.style.use('dark_background')
            fig, ax1 = plt.subplots(figsize=(14, 8))
            fig.patch.set_facecolor(COLORS['surface'])
            
            # Agrupar por PlanType y ordenar descendente
            plantype_data = self.slowmotion_data.groupby('PlanType')['ExtOH'].sum().sort_values(ascending=False)
            
            if len(plantype_data) > 0:
                # Calcular porcentajes acumulativos
                total_value = plantype_data.sum()
                cumulative_pct = (plantype_data.cumsum() / total_value * 100)
                
                # Filtrar para mostrar solo hasta el 80%
                pareto_mask = cumulative_pct <= 80
                # Incluir el primer elemento que supere el 80% para completar la vista
                if not pareto_mask.all():
                    first_over_80 = cumulative_pct[cumulative_pct > 80].index[0]
                    pareto_mask[first_over_80] = True
                
                # Datos filtrados para Pareto
                pareto_data = plantype_data[pareto_mask]
                pareto_cumulative = cumulative_pct[pareto_mask]
                
                # Crear eje secundario para porcentajes
                ax2 = ax1.twinx()
                
                # Gráfico de barras (valores absolutos)
                colors = ['#4ECDC4', '#45B7D1', '#FF6B6B', '#FFA726', '#AB47BC', '#26C6DA', '#66BB6A', '#EC407A']
                bars = ax1.bar(range(len(pareto_data)), pareto_data.values,
                              color=colors[:len(pareto_data)], alpha=0.8, 
                              edgecolor='white', linewidth=1.5)
                
                # Línea de porcentaje acumulativo
                line = ax2.plot(range(len(pareto_data)), pareto_cumulative.values,
                               color='#FFD700', marker='o', linewidth=3, markersize=8,
                               label='% Acumulativo')
                
                # Línea del 80%
                ax2.axhline(y=80, color='#FF6B6B', linestyle='--', linewidth=2, 
                           alpha=0.8, label='80% Target')
                
                # Agregar valores en las barras
                for i, (bar, value, pct) in enumerate(zip(bars, pareto_data.values, pareto_cumulative.values)):
                    height = bar.get_height()
                    # Valor absoluto en la barra
                    ax1.text(bar.get_x() + bar.get_width()/2., height + max(pareto_data.values)*0.01,
                            f'${value:,.0f}', ha='center', va='bottom',
                            color='white', fontsize=11, fontweight='bold')
                    
                    # Porcentaje acumulativo en la línea
                    ax2.text(i, pct + 2, f'{pct:.1f}%', ha='center', va='bottom',
                            color='#FFD700', fontsize=10, fontweight='bold')
                
                # Configurar eje primario (barras)
                ax1.set_xticks(range(len(pareto_data)))
                ax1.set_xticklabels(pareto_data.index, rotation=45, ha='right', fontsize=12)
                ax1.set_title('Diagrama de Pareto - Inventario Obsoleto por Plan Type (80%)',
                             color='white', fontsize=18, fontweight='bold', pad=20)
                ax1.set_ylabel('Valor Total (USD)', color='white', fontsize=14, fontweight='bold')
                ax1.grid(True, alpha=0.3, color='white', axis='y')
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
                
                # Configurar eje secundario (línea)
                ax2.set_ylabel('Porcentaje Acumulativo (%)', color='#FFD700', fontsize=14, fontweight='bold')
                ax2.set_ylim(0, 100)
                ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:.0f}%'))
                
                # Leyenda
                ax2.legend(loc='center right', fontsize=12, fancybox=True, shadow=True)
                
                # Estilo de los ejes
                ax1.set_facecolor('#2C3E50')
                ax1.tick_params(colors='white')
                ax2.tick_params(axis='y', colors='#FFD700')
                
                for spine in ax1.spines.values():
                    spine.set_color('white')
                for spine in ax2.spines.values():
                    spine.set_color('#FFD700' if spine.spine_type == 'right' else 'white')
                
                # Agregar información del Pareto
                total_items_shown = len(pareto_data)
                total_items = len(plantype_data)
                value_shown = pareto_data.sum()
                pct_value_shown = (value_shown / total_value) * 100
                
                info_text = f'Mostrando {total_items_shown}/{total_items} PlanTypes • {pct_value_shown:.1f}% del valor total'
                fig.text(0.5, 0.02, info_text, ha='center', va='bottom', 
                        color='white', fontsize=12, style='italic')
                
            else:
                ax1.text(0.5, 0.5, 'No hay datos de PlanType', ha='center', va='center',
                        color='white', fontsize=16, transform=ax1.transAxes)
            
            plt.tight_layout()
            plt.subplots_adjust(bottom=0.15)  # Espacio para el texto informativo
            
            # Convertir a base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', facecolor=COLORS['surface'],
                       bbox_inches='tight', dpi=150)
            buffer.seek(0)
            chart_base64 = base64.b64encode(buffer.getvalue()).decode()
            plt.close()
            
            return chart_base64
            
        except Exception as e:
            print(f"Error creando gráfico Pareto: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def update_obsolete_items_table(self, db_path=r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"):
        """Actualizar tabla con nueva lógica: ItemNo + Lot + QtyOH_SUM"""
        try:
            print("🔄 Actualizando tabla de items obsoletos...")
            conn = sqlite3.connect(db_path)
            
            if self.slowmotion_data is None:
                print("❌ No hay datos de slowmotion para actualizar")
                return {"success": False, "message": "No hay datos de slowmotion"}
            
            print(f"📊 Slowmotion original: {len(self.slowmotion_data)} registros")
            
            # NUEVA LÓGICA: AGRUPAR POR ItemNo + Lot y SUMAR QtyOH
            print("🔧 Agrupando por ItemNo + Lot y sumando QtyOH...")
            slowmotion_grouped = self.slowmotion_data.groupby(['ItemNo', 'Lot']).agg({
                'Description': 'first',  # Tomar primera descripción
                'UM': 'first',
                'S': 'first', 
                'WH': 'first',
                'LastDate': 'first',
                'Bin': lambda x: ','.join(x.astype(str)),  # Concatenar bins
                'Rev': 'first',
                'Entity': 'first',
                'PlanType': 'first',
                'Prj': 'first',
                'PrdGrp': 'first',
                'MLI': 'first',
                'LT': 'first',
                'PlnrCd': 'first',
                'ABCD': 'first',
                'BM': 'first',
                'Cost': 'mean',  # Promedio de costos
                'QtyOH': 'sum',  # SUMAR CANTIDAD
                'ExtOH': 'sum'   # SUMAR VALOR EXTENDIDO
            }).reset_index()
            
            print(f"📊 Después de agrupar: {len(slowmotion_grouped)} registros únicos")
            
            # VERIFICAR SI LA TABLA ESTÁ VACÍA
            count_query = "SELECT COUNT(*) as count FROM newObsolet_slowmotion_items"
            count_result = pd.read_sql_query(count_query, conn)
            existing_count = count_result['count'].iloc[0]
            
            if existing_count == 0:
                print("📋 Tabla vacía, inicializando con datos agrupados...")
                
                # BORRAR REGISTROS EXISTENTES
                conn.execute("DELETE FROM newObsolet_slowmotion_items")
                conn.commit()
                
                # Preparar datos para inserción
                init_data = slowmotion_grouped.copy()
                init_data['date_added'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                init_data['root_cause'] = ''
                init_data['status'] = 'Historical'
                init_data['engineer_notes'] = ''
                init_data['is_new'] = 0
                
                # INSERTAR EN LOTES
                batch_size = 500
                total_rows = len(init_data)
                inserted_count = 0
                
                print(f"📊 Insertando {total_rows} registros agrupados en lotes de {batch_size}...")
                
                for i in range(0, total_rows, batch_size):
                    batch = init_data.iloc[i:i+batch_size]
                    try:
                        batch.to_sql('newObsolet_slowmotion_items', conn, if_exists='append', index=False)
                        inserted_count += len(batch)
                        progress = (inserted_count / total_rows) * 100
                        print(f"✅ Lote {i//batch_size + 1}: {inserted_count}/{total_rows} ({progress:.1f}%)")
                    except Exception as batch_error:
                        print(f"❌ Error en lote {i//batch_size + 1}: {batch_error}")
                
                print(f"✅ Inserción completada: {inserted_count}/{total_rows} registros")
                
                # Recargar datos
                obsolete_query = "SELECT * FROM newObsolet_slowmotion_items"
                self.obsolete_items_data = pd.read_sql_query(obsolete_query, conn)
                
                conn.close()
                self._calculate_metrics()
                
                return {
                    "success": True, 
                    "message": f"Tabla inicializada con {inserted_count} items agrupados",
                    "new_items_count": 0
                }
            
            # LÓGICA PARA ITEMS NUEVOS - NUEVA COMPARACIÓN
            print("🔍 Comparando items existentes vs slowmotion agrupado...")
            
            # Obtener items existentes agrupados
            existing_query = """
            SELECT ItemNo, Lot, QtyOH FROM newObsolet_slowmotion_items
            """
            existing_items = pd.read_sql_query(existing_query, conn)
            
            # Crear llaves de comparación: ItemNo + Lot + QtyOH
            existing_keys = set(
                existing_items['ItemNo'] + '|' + 
                existing_items['Lot'].astype(str) + '|' + 
                existing_items['QtyOH'].astype(str)
            )
            
            slowmotion_keys = (
                slowmotion_grouped['ItemNo'] + '|' + 
                slowmotion_grouped['Lot'].astype(str) + '|' + 
                slowmotion_grouped['QtyOH'].astype(str)
            )
            
            # Identificar nuevos items
            new_mask = ~slowmotion_keys.isin(existing_keys)
            new_items = slowmotion_grouped[new_mask].copy()
            
            print(f"🆕 Items nuevos detectados: {len(new_items)}")
            
            if len(new_items) > 0:
                # Agregar columnas adicionales
                new_items['date_added'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                new_items['root_cause'] = ''
                new_items['status'] = 'Pending Analysis'
                new_items['engineer_notes'] = ''
                new_items['is_new'] = 1
                
                # Insertar nuevos items
                if len(new_items) > 500:
                    print(f"📊 Insertando {len(new_items)} nuevos items en lotes...")
                    batch_size = 500
                    for i in range(0, len(new_items), batch_size):
                        batch = new_items.iloc[i:i+batch_size]
                        batch.to_sql('newObsolet_slowmotion_items', conn, if_exists='append', index=False)
                        print(f"✅ Lote nuevo {i//batch_size + 1} insertado")
                else:
                    new_items.to_sql('newObsolet_slowmotion_items', conn, if_exists='append', index=False)
                
                print(f"✅ Se agregaron {len(new_items)} nuevos items obsoletos")
                
                # Recargar datos
                obsolete_query = "SELECT * FROM newObsolet_slowmotion_items"
                self.obsolete_items_data = pd.read_sql_query(obsolete_query, conn)
                
                # Actualizar items nuevos
                self.new_items = self.obsolete_items_data[
                    self.obsolete_items_data['is_new'] == 1
                ].copy()
                
                conn.close()
                self._calculate_metrics()
                
                return {
                    "success": True, 
                    "message": f"Se agregaron {len(new_items)} nuevos items",
                    "new_items_count": len(new_items)
                }
            else:
                print("✅ No hay nuevos items para agregar")
                conn.close()
                return {
                    "success": True, 
                    "message": "No hay nuevos items para agregar",
                    "new_items_count": 0
                }
                
        except Exception as e:
            print(f"❌ Error actualizando tabla: {e}")
            import traceback
            traceback.print_exc()
            return {"success": False, "message": f"Error: {str(e)}"}

    def update_item_root_cause(self, item_id, root_cause, engineer_notes, 
                            db_path=r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"):
        """Actualizar causa raíz de un item"""
        try:
            conn = sqlite3.connect(db_path)
            
            update_query = """
            UPDATE newObsolet_slowmotion_items 
            SET root_cause = ?, engineer_notes = ?, status = 'Analyzed', is_new = 0
            WHERE id = ?
            """
            
            conn.execute(update_query, (root_cause, engineer_notes, item_id))
            conn.commit()
            conn.close()
            
            # Recargar datos
            self.load_data_from_db(db_path)
            
            return {"success": True, "message": "Item actualizado exitosamente"}
            
        except Exception as e:
            return {"success": False, "message": f"Error: {str(e)}"}

    def create_department_chart(self):
        """Crear gráfico de distribución por departamentos"""
        if self.obsolete_items_data is None or len(self.obsolete_items_data) == 0:
            return None
        
        try:
            import matplotlib
            matplotlib.use('Agg')
            
            plt.style.use('dark_background')
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
            fig.patch.set_facecolor(COLORS['surface'])
            
            # Gráfico 1: Distribución por Departamento
            dept_data = self.obsolete_items_data.groupby('department').agg({
                'ExtOH': 'sum',
                'ItemNo': 'count'
            }).fillna({'department': 'Sin Asignar'})
            
            if len(dept_data) > 0:
                colors = ['#4ECDC4', '#45B7D1', '#FF6B6B', '#FFA726', '#AB47BC', '#26C6DA', '#66BB6A']
                
                # Pie chart de departamentos
                wedges, texts, autotexts = ax1.pie(
                    dept_data['ExtOH'].values, 
                    labels=dept_data.index,
                    colors=colors[:len(dept_data)], 
                    autopct='%1.1f%%',
                    startangle=90,
                    textprops={'color': 'white', 'fontsize': 10}
                )
                
                ax1.set_title('Distribución de Valor por Departamento', 
                            color='white', fontsize=16, fontweight='bold', pad=20)
            
            # Gráfico 2: Status por Departamento
            status_dept = self.obsolete_items_data.groupby(['department', 'status']).size().unstack(fill_value=0)
            
            if not status_dept.empty:
                status_dept.plot(kind='bar', ax=ax2, color=['#FF6B6B', '#FFA726', '#4ECDC4', '#45B7D1'])
                ax2.set_title('Status por Departamento', color='white', fontsize=16, fontweight='bold')
                ax2.set_xlabel('Departamento', color='white', fontsize=14)
                ax2.set_ylabel('Cantidad de Items', color='white', fontsize=14)
                ax2.tick_params(colors='white')
                ax2.legend(loc='upper right')
            
            # Estilo general
            for ax in [ax1, ax2]:
                ax.set_facecolor('#2C3E50')
                for spine in ax.spines.values():
                    spine.set_color('white')
            
            plt.tight_layout()
            
            # Convertir a base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', facecolor=COLORS['surface'],
                    bbox_inches='tight', dpi=150)
            buffer.seek(0)
            chart_base64 = base64.b64encode(buffer.getvalue()).decode()
            plt.close()
            
            return chart_base64
            
        except Exception as e:
            print(f"Error creando gráfico departamentos: {e}")
            return None

    def update_item_assignment(self, item_id, department, responsible_person, 
                            db_path=r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"):
        """Actualizar asignación de departamento y responsable"""
        try:
            conn = sqlite3.connect(db_path)
            
            update_query = """
            UPDATE newObsolet_slowmotion_items 
            SET department = ?, responsible_person = ?, assigned_date = ?, status = 'Assigned'
            WHERE id = ?
            """
            
            assigned_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            conn.execute(update_query, (department, responsible_person, assigned_date, item_id))
            conn.commit()
            conn.close()
            
            # Recargar datos
            self.load_data_from_db(db_path)
            
            return {"success": True, "message": "Asignación actualizada exitosamente"}
            
        except Exception as e:
            return {"success": False, "message": f"Error: {str(e)}"}


    def _create_department_tab(self):
        """Crear tab de gestión por departamentos"""
        
        # Panel de control para departamentos
        department_control = ft.Container(
            content=ft.Column([
                ft.Text("🏢 Gestión por Departamentos", size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Row([
                    ft.ElevatedButton(
                        text="📊 Actualizar Análisis",
                        on_click=self._refresh_department_analysis,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['accent'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=200
                    ),
                    ft.ElevatedButton(
                        text="📊 Exportar por Depto",
                        on_click=self._export_by_department,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['success'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=200
                    )
                ], spacing=15)
            ], spacing=15),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Gráfico de departamentos
        self.department_chart_container = ft.Container(
            content=ft.Text("📈 El gráfico de departamentos aparecerá aquí",
                        size=16, color=COLORS['text_secondary']),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Filtros
        self.department_filters = ft.Container(
            content=ft.Row([
                ft.Dropdown(
                    label="Filtrar por Departamento",
                    options=[
                        ft.dropdown.Option("", "Todos"),
                        ft.dropdown.Option("PRODUCCION", "Producción"),
                        ft.dropdown.Option("ALMACEN", "Almacén"),
                        ft.dropdown.Option("CALIDAD", "Calidad"),
                        ft.dropdown.Option("INGENIERIA", "Ingeniería"),
                        ft.dropdown.Option("COMPRAS", "Compras"),
                        ft.dropdown.Option("VENTAS", "Ventas"),
                        ft.dropdown.Option("MANTENIMIENTO", "Mantenimiento")
                    ],
                    width=200,
                    on_change=self._filter_by_department
                ),
                ft.Dropdown(
                    label="Filtrar por Status",
                    options=[
                        ft.dropdown.Option("", "Todos"),
                        ft.dropdown.Option("Pending Analysis", "Pendiente Análisis"),
                        ft.dropdown.Option("Analyzed", "Analizado"),
                        ft.dropdown.Option("Assigned", "Asignado")
                    ],
                    width=200,
                    on_change=self._filter_by_status
                )
            ], spacing=15),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Tabla de asignaciones
        self.department_table = ft.DataTable(
            columns=[
                ft.DataColumn(label=ft.Text("Item No", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Descripción", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Valor ExtOH", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Status", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Departamento", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Responsable", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Acciones", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']))
            ],
            rows=[],
            bgcolor=COLORS['secondary'],
            border_radius=12
        )
        
        self.department_table_container = ft.Container(
            content=ft.Column([
                ft.Text("📋 Asignación por Departamentos", size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Container(
                    content=ft.Column([
                        self.department_table
                    ], scroll=ft.ScrollMode.AUTO),
                    bgcolor=COLORS['secondary'],
                    border_radius=12,
                    padding=10,
                    height=400
                )
            ]),
            visible=False
        )
        
        return ft.Column([
            department_control,
            self.department_chart_container,
            self.department_filters,
            self.department_table_container
        ], scroll=ft.ScrollMode.AUTO)

    async def _refresh_department_analysis(self, e):
        """Actualizar análisis de departamentos"""
        await self._show_loading()
        try:
            await self._update_department_chart()
            await self._update_department_table()
            await self._show_snackbar("✅ Análisis por departamentos actualizado")
        except Exception as ex:
            await self._show_snackbar(f"❌ Error: {str(ex)}")
        await self._hide_loading()

    async def _update_department_chart(self):
        """Actualizar gráfico de departamentos"""
        chart_base64 = self.analyzer.create_department_chart()
        
        if chart_base64:
            chart_image = ft.Image(
                src_base64=chart_base64,
                width=1200,
                height=600,
                fit=ft.ImageFit.CONTAIN
            )
            
            self.department_chart_container.content = ft.Column([
                ft.Text("📈 Análisis por Departamentos", 
                    size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                chart_image
            ], spacing=15)
        else:
            self.department_chart_container.content = ft.Text(
                "❌ Error generando gráfico de departamentos",
                size=16, color=COLORS['error']
            )
        
        self.page.update()

    async def _update_department_table(self):
        """Actualizar tabla de departamentos"""
        if self.analyzer.obsolete_items_data is None or len(self.analyzer.obsolete_items_data) == 0:
            self.department_table_container.visible = False
            return
        
        # Crear filas con opciones de asignación
        rows = []
        for _, item in self.analyzer.obsolete_items_data.head(50).iterrows():  # Limitar a 50 para performance
            # Dropdown para departamento
            dept_dropdown = ft.Dropdown(
                value=item.get('department', ''),
                options=[
                    ft.dropdown.Option("", "Sin Asignar"),
                    ft.dropdown.Option("PRODUCCION", "Producción"),
                    ft.dropdown.Option("ALMACEN", "Almacén"),
                    ft.dropdown.Option("CALIDAD", "Calidad"),
                    ft.dropdown.Option("INGENIERIA", "Ingeniería"),
                    ft.dropdown.Option("COMPRAS", "Compras"),
                    ft.dropdown.Option("VENTAS", "Ventas"),
                    ft.dropdown.Option("MANTENIMIENTO", "Mantenimiento")
                ],
                width=130,
                text_size=10
            )
            
            # Campo para responsable
            responsible_field = ft.TextField(
                value=str(item.get('responsible_person', '')),
                hint_text="Responsable...",
                width=120,
                height=35,
                text_size=10
            )
            
            # Botón de guardar
            save_btn = ft.ElevatedButton(
                text="💾",
                tooltip="Guardar asignación",
                on_click=lambda e, item_id=item.get('id'), dept=dept_dropdown, resp=responsible_field: 
                    self._save_assignment(item_id, dept, resp),
                style=ft.ButtonStyle(bgcolor=COLORS['success']),
                width=40,
                height=30
            )
            
            row = ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(str(item.get('ItemNo', ''))[:15], size=10, color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(str(item.get('Description', ''))[:20], size=10, color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(f"${float(item.get('ExtOH', 0)):,.0f}", size=10, color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(str(item.get('status', ''))[:10], size=10, color=COLORS['text_primary'])),
                    ft.DataCell(dept_dropdown),
                    ft.DataCell(responsible_field),
                    ft.DataCell(save_btn)
                ]
            )
            rows.append(row)
        
        self.department_table.rows = rows
        self.department_table_container.visible = True
        self.page.update()

    async def _save_assignment(self, item_id, dept_dropdown, responsible_field):
        """Guardar asignación de departamento y responsable"""
        try:
            department = dept_dropdown.value
            responsible = responsible_field.value.strip()
            
            result = self.analyzer.update_item_assignment(item_id, department, responsible)
            
            if result["success"]:
                await self._show_snackbar("✅ Asignación guardada exitosamente")
                await self._refresh_department_analysis(None)
            else:
                await self._show_snackbar(f"❌ {result['message']}")
                
        except Exception as ex:
            await self._show_snackbar(f"❌ Error: {str(ex)}")

    async def _export_by_department(self, e):
        """Exportar por departamento"""
        await self._show_snackbar("🚧 Función de exportación por departamento en desarrollo")

    async def _filter_by_department(self, e):
        """Filtrar por departamento"""
        await self._show_snackbar("🚧 Filtro por departamento en desarrollo")

    async def _filter_by_status(self, e):
        """Filtrar por status"""
        await self._show_snackbar("🚧 Filtro por status en desarrollo")




class ObsoleteAnalyzerApp:
    def __init__(self):
        self.analyzer = ObsoleteSlowMotionAnalyzer()
        self.page = None
        self.loading_modal = None
        
    def main(self, page: ft.Page):
        """Función principal de la aplicación"""
        self.page = page
        page.title = "🏭 Obsolete Slow Motion Items Analyzer"
        page.theme_mode = ft.ThemeMode.DARK
        page.bgcolor = COLORS['surface']
        page.window_width = 1400
        page.window_height = 900
        page.window_resizable = True
        page.scroll = ft.ScrollMode.AUTO
        
        # Modal de carga
        self.loading_modal = ft.AlertDialog(
            modal=True,
            content=ft.Column([
                ft.ProgressRing(width=50, height=50, color=COLORS['accent']),
                ft.Text("Procesando datos...", size=16, color=COLORS['text_primary'])
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, height=100),
            bgcolor=COLORS['secondary']
        )
        
        # Header principal
        header = self._create_header()
        
        # Crear tabs
        # Crear tabs CON NUEVO TAB
        tabs = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=[
                ft.Tab(
                    text="📊 Análisis por PlanType",
                    icon=ft.Icons.BAR_CHART,
                    content=self._create_analysis_tab()
                ),
                ft.Tab(
                    text="🔧 Gestión de Items Obsoletos",
                    icon=ft.Icons.ENGINEERING,
                    content=self._create_management_tab()
                ),
                ft.Tab(
                    text="🏢 Gestión por Departamentos",
                    icon=ft.Icons.BUSINESS_CENTER,
                    content=self._create_department_tab()
                )
            ],
            expand=True
        )
        
        # Layout principal
        main_content = ft.Column([
            header,
            tabs
        ], expand=True)
        
        page.add(main_content)
        page.update()
    
    def _create_header(self):
        """Crear header de la aplicación"""
        return ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.SLOW_MOTION_VIDEO, size=40, color=COLORS['accent']),
                ft.Column([
                    ft.Text("Obsolete Slow Motion Items Analyzer", size=28, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Análisis y gestión de items de lento movimiento", size=14, color=COLORS['text_secondary'])
                ], spacing=0),
                ft.Container(expand=True),
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.ACCESS_TIME, size=16, color=COLORS['text_secondary']),
                        ft.Text(f"Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                               size=12, color=COLORS['text_secondary'])
                    ], spacing=5),
                    alignment=ft.alignment.center_right
                )
            ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=ft.border_radius.only(top_left=12, top_right=12),
            margin=ft.margin.only(bottom=20)
        )
    
    def _create_analysis_tab(self):
        """Crear tab de análisis"""
        # Panel de control para análisis
        control_panel = ft.Container(
            content=ft.Column([
                ft.Text("🎛️ Panel de Control", size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Row([
                    ft.ElevatedButton(
                        text="🗄️ Cargar desde R4Database",
                        on_click=self._load_from_db,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['accent'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=250
                    ),
                    ft.ElevatedButton(
                        text="🧪 Usar Datos de Ejemplo",
                        on_click=self._load_sample_data,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['warning'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=250
                    )
                ], spacing=15)
            ], spacing=15),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Contenedor de métricas
        self.metrics_container = ft.Container(
            content=ft.Text("📊 Carga datos para ver métricas",
                          size=16, color=COLORS['text_secondary']),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Contenedor de gráficos
        self.chart_container = ft.Container(
            content=ft.Text("📈 El gráfico aparecerá aquí después del análisis",
                          size=16, color=COLORS['text_secondary']),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        return ft.Column([
            control_panel,
            self.metrics_container,
            self.chart_container
        ], scroll=ft.ScrollMode.AUTO)
    
    def _create_management_tab(self):
        """Crear tab de gestión de items obsoletos"""
        # Panel de control para gestión
        management_control = ft.Container(
            content=ft.Column([
                ft.Text("🔧 Gestión de Items Obsoletos", size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Row([
                    ft.ElevatedButton(
                        text="🔄 Actualizar Items Obsoletos",
                        on_click=self._update_obsolete_items,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['success'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=280
                    ),
                    ft.ElevatedButton(
                        text="📊 Exportar Reporte",
                        on_click=self._export_obsolete_report,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['accent'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=280
                    )
                ], spacing=15)
            ], spacing=15),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Status de nuevos items
        self.new_items_status = ft.Container(
            content=ft.Text("🔍 Carga datos para ver items nuevos",
                          size=16, color=COLORS['text_secondary']),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Tabla de items nuevos pendientes
        self.new_items_table = ft.DataTable(
            columns=[
                ft.DataColumn(label=ft.Text("Item No", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Descripción", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("PlanType", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Valor ExtOH", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Fecha Agregado", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Causa Raíz", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Acciones", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']))
            ],
            rows=[],
            bgcolor=COLORS['secondary'],
            border_radius=12
        )
        
        self.table_container = ft.Container(
            content=ft.Column([
                ft.Text("📋 Items Nuevos Pendientes de Análisis", size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Container(
                    content=self.new_items_table,
                    bgcolor=COLORS['secondary'],
                    border_radius=12,
                    padding=10
                )
            ]),
            visible=False
        )
        
        return ft.Column([
            management_control,
            self.new_items_status,
            self.table_container
        ], scroll=ft.ScrollMode.AUTO)
    
    async def _load_from_db(self, e):
        """Cargar datos desde R4Database"""
        await self._show_loading()
        
        try:
            success = self.analyzer.load_data_from_db()
            if success:
                await self._update_analysis_ui("✅ Datos cargados desde R4Database exitosamente")
            else:
                await self._update_analysis_ui("⚠️ Error conectando a DB, usando datos de ejemplo")
        except Exception as ex:
            await self._update_analysis_ui(f"❌ Error: {str(ex)}")
        
        await self._hide_loading()
    
    async def _load_sample_data(self, e):
        """Cargar datos de ejemplo"""
        await self._show_loading()
        
        try:
            self.analyzer._load_sample_data()
            await self._update_analysis_ui("✅ Datos de ejemplo cargados exitosamente")
        except Exception as ex:
            await self._update_analysis_ui(f"❌ Error: {str(ex)}")
        
        await self._hide_loading()
    
    async def _update_obsolete_items(self, e):
        """Actualizar tabla de items obsoletos"""
        await self._show_loading()
        
        try:
            result = self.analyzer.update_obsolete_items_table()
            if result["success"]:
                await self._update_management_ui(f"✅ {result['message']}")
                await self._show_snackbar(f"Se encontraron {result['new_items_count']} nuevos items")
            else:
                await self._show_snackbar(f"❌ {result['message']}")
        except Exception as ex:
            await self._show_snackbar(f"❌ Error: {str(ex)}")
        
        await self._hide_loading()
    
    async def _export_obsolete_report(self, e):
        """Exportar reporte de items obsoletos"""
        try:
            if self.analyzer.obsolete_items_data is None or len(self.analyzer.obsolete_items_data) == 0:
                await self._show_snackbar("❌ No hay datos para exportar")
                return
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            export_path = f"Obsolete_Items_Report_{timestamp}.xlsx"
            
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                # Todos los items obsoletos
                self.analyzer.obsolete_items_data.to_excel(writer, sheet_name='All_Obsolete_Items', index=False)
                
                # Items nuevos pendientes
                if self.analyzer.new_items is not None and len(self.analyzer.new_items) > 0:
                    self.analyzer.new_items.to_excel(writer, sheet_name='New_Items_Pending', index=False)
                
                # Análisis por PlanType
                if 'plantype_analysis' in self.analyzer.metrics:
                    self.analyzer.metrics['plantype_analysis'].to_excel(writer, sheet_name='PlanType_Analysis', index=False)
                
                # Datos originales de slowmotion
                if self.analyzer.slowmotion_data is not None:
                    self.analyzer.slowmotion_data.to_excel(writer, sheet_name='Slowmotion_Data', index=False)
                
                # Resumen de métricas
                summary_data = {
                    'Métrica': [
                        'Total Items Slowmotion',
                        'Valor Total Obsoleto',
                        'Items Nuevos Pendientes',
                        'Items Pending Analysis',
                        'Días Promedio en Sistema',
                        'Top Entity',
                        'Top Project',
                        'Unique Entities',
                        'Unique Projects'
                    ],
                    'Valor': [
                        self.analyzer.metrics.get('total_items', 0),
                        f"${self.analyzer.metrics.get('total_value', 0):,.2f}",
                        self.analyzer.metrics.get('new_items_count', 0),
                        self.analyzer.metrics.get('pending_analysis_count', 0),
                        f"{self.analyzer.metrics.get('avg_days_in_system', 0)} días",
                        self.analyzer.metrics.get('top_entity', 'N/A'),
                        self.analyzer.metrics.get('top_project', 'N/A'),
                        self.analyzer.metrics.get('unique_entities', 0),
                        self.analyzer.metrics.get('unique_projects', 0)
                    ]
                }
                pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary_Metrics', index=False)
            
            # Abrir carpeta donde se guardó el archivo
            full_path = os.path.abspath(export_path)
            if platform.system() == "Windows":
                subprocess.Popen(f'explorer /select,"{full_path}"')
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", "-R", full_path])
            else:  # Linux
                subprocess.Popen(["xdg-open", os.path.dirname(full_path)])
            
            await self._show_snackbar(f"✅ Reporte exportado: {os.path.basename(export_path)}")
            
        except Exception as ex:
            await self._show_snackbar(f"❌ Error exportando: {str(ex)}")
    
    async def _show_loading(self):
        """Mostrar modal de carga"""
        self.page.dialog = self.loading_modal
        self.loading_modal.open = True
        self.page.update()
        await asyncio.sleep(0.1)
    
    async def _hide_loading(self):
        """Ocultar modal de carga"""
        if self.loading_modal:
            self.loading_modal.open = False
        self.page.update()
    
    async def _show_snackbar(self, message):
        """Mostrar mensaje temporal"""
        snackbar = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE),
            bgcolor=COLORS['primary'],
            duration=3000
        )
        self.page.snack_bar = snackbar
        snackbar.open = True
        self.page.update()
    
    async def _update_analysis_ui(self, status_message):
        """Actualizar UI del tab de análisis"""
        if self.analyzer.slowmotion_data is not None and len(self.analyzer.slowmotion_data) > 0:
            # Actualizar métricas
            await self._update_metrics()
            
            # Actualizar gráfico
            await self._update_chart()
            
            await self._show_snackbar(status_message)
        else:
            await self._show_snackbar("❌ No se pudieron cargar los datos")
    
    async def _update_management_ui(self, status_message):
        """Actualizar UI del tab de gestión"""
        if self.analyzer.obsolete_items_data is not None:
            # Actualizar status de nuevos items
            await self._update_new_items_status()
            
            # Actualizar tabla de items nuevos
            await self._update_new_items_table()
            
            await self._show_snackbar(status_message)

    def _create_enhanced_metric_card(self, icon, title, value, subtitle, color, icon_color):
        """Crear KPI card con estilo Credit Memo - bordes de color y fondo transparente"""
        return ft.Container(
            content=ft.Column([
                # Icono pequeño en la esquina superior izquierda
                ft.Row([
                    ft.Container(
                        content=ft.Icon(icon, size=20, color=color),  # TAMAÑO ICONO
                        margin=ft.margin.only(bottom=10)
                    ),
                    ft.Container(expand=True),
                    ft.Text(title.upper(), size=11, color="#bdbdbd", weight=ft.FontWeight.BOLD)  # TÍTULO
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                
                # Valor principal muy grande y centrado
                ft.Container(
                    content=ft.Text(
                        value,
                        size=42,  # TAMAÑO VALOR PRINCIPAL
                        weight=ft.FontWeight.BOLD,
                        color="#ffffff",  # COLOR VALOR
                        text_align=ft.TextAlign.CENTER
                    ),
                    alignment=ft.alignment.center,
                    margin=ft.margin.symmetric(vertical=15)  # ESPACIADO VERTICAL
                ),
                
                # Subtítulo centrado
                ft.Text(
                    subtitle,
                    size=12,  # TAMAÑO SUBTÍTULO
                    color="#757575",  # COLOR SUBTÍTULO
                    text_align=ft.TextAlign.CENTER,
                    weight=ft.FontWeight.W_400
                )
            ], 
            spacing=0,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
            ),
            width=280,  # ANCHO CARD
            height=160,  # ALTO CARD
            padding=20,  # PADDING INTERNO
            bgcolor="#1a1a1a",  # COLOR FONDO
            border=ft.border.all(2, color),  # GROSOR Y COLOR BORDE
            border_radius=15  # RADIO ESQUINAS
        )

    def _create_info_metric_card(self, icon, title, value, additional_info, color):
        """Crear card de información secundaria con estilo Credit Memo"""
        return ft.Container(
            content=ft.Column([
                # Header con icono y título
                ft.Row([
                    ft.Container(
                        content=ft.Icon(icon, size=18, color="#00bcd4"),  # TAMAÑO ICONO
                        margin=ft.margin.only(right=8)
                    ),
                    ft.Text(title.upper(), size=11, color="#bdbdbd", weight=ft.FontWeight.BOLD)  # TÍTULO
                ], alignment=ft.MainAxisAlignment.START),
                
                # Valor principal
                ft.Container(
                    content=ft.Text(
                        value,
                        size=24,  # TAMAÑO VALOR
                        weight=ft.FontWeight.BOLD,
                        color="#ffffff",  # COLOR VALOR
                        text_align=ft.TextAlign.CENTER
                    ),
                    alignment=ft.alignment.center,
                    margin=ft.margin.symmetric(vertical=8)  # ESPACIADO
                ),
                
                # Subtítulo
                ft.Text(
                    additional_info,  # Nota: aquí era 'additional_info', no 'subtitle'
                    size=11,  # TAMAÑO SUBTÍTULO
                    color="#757575",  # COLOR SUBTÍTULO
                    text_align=ft.TextAlign.CENTER
                )
            ], 
            spacing=5,  # ESPACIADO ENTRE ELEMENTOS
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
            ),
            width=320,  # ANCHO CARD
            height=120,  # ALTO CARD
            padding=20,  # PADDING INTERNO
            bgcolor="#1a1a1a",  # COLOR FONDO
            border=ft.border.all(2, "#00bcd4"),  # GROSOR Y COLOR BORDE
            border_radius=15  # RADIO ESQUINAS
        )

    async def _update_metrics(self):
        """Actualizar panel de métricas con estilo moderno"""
        if not self.analyzer.metrics:
            return
        
        m = self.analyzer.metrics
        
        # Cards principales con el nuevo estilo (4 cards superiores)
        metrics_row1 = ft.Row([
            self._create_enhanced_metric_card(
                icon=ft.Icons.INVENTORY_2,
                title="Total Items",
                value=f"{m.get('total_items', 0):,}",
                subtitle="Items en slowmotion",
                color="#00bcd4",  # Azul
                icon_color="#00bcd4"  # Mismo color
            ),
            self._create_enhanced_metric_card(
                icon=ft.Icons.ATTACH_MONEY,
                title="Valor Total",
                value=f"${m.get('total_value', 0)/1000:,.0f}K",
                subtitle=f"${m.get('total_value', 0):,.0f} USD",
                color="#4caf50",  # Verde
                icon_color="#4caf50"
            ),
            self._create_enhanced_metric_card(
                icon=ft.Icons.NEW_RELEASES,
                title="Items Nuevos",
                value=f"{m.get('new_items_count', 0):,}",
                subtitle="Pendientes análisis",
                color="#ff9800",  # Naranja
                icon_color="#ff9800"
            ),
            self._create_enhanced_metric_card(
                icon=ft.Icons.PENDING_ACTIONS,
                title="Pendientes",
                value=f"{m.get('pending_analysis_count', 0):,}",
                subtitle="Requieren causa raíz",
                color="#f44336",  # Rojo
                icon_color="#f44336"
            )
        ], spacing=20, wrap=True)
        
        # Cards de información adicional (fila inferior)
        metrics_row2 = ft.Row([
            self._create_info_metric_card(
                icon=ft.Icons.BUSINESS,
                title="Top Entity",
                value=f"{m.get('top_entity', 'N/A')}",
                additional_info=f"Total: {m.get('unique_entities', 0)} entities",
                color="#37474F"
            ),
            self._create_info_metric_card(
                icon=ft.Icons.FOLDER_SPECIAL,
                title="Top Project",
                value=f"{m.get('top_project', 'N/A')}",
                additional_info=f"Total: {m.get('unique_projects', 0)} projects",
                color="#37474F"
            ),
            self._create_info_metric_card(
                icon=ft.Icons.SCHEDULE,
                title="Días Promedio",
                value=f"{m.get('avg_days_in_system', 0):.0f}",
                additional_info="En el sistema",
                color="#37474F"
            )
        ], spacing=20, wrap=True)
        
        self.metrics_container.content = ft.Column([
            ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.DASHBOARD, size=24, color="#00bcd4"),
                    ft.Text("Métricas de Items Obsoletos", size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])
                ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
                margin=ft.margin.only(bottom=20)
            ),
            metrics_row1,
            ft.Container(height=15),  # Espaciado entre filas
            metrics_row2
        ], spacing=0)
        
        self.page.update()

    async def _update_chart(self):
        """Actualizar gráfico de PlanType"""
        chart_base64 = self.analyzer.create_plantype_chart()
        
        if chart_base64:
            chart_image = ft.Image(
                src_base64=chart_base64,
                width=1200,
                height=600,
                fit=ft.ImageFit.CONTAIN
            )
            
            self.chart_container.content = ft.Column([
                ft.Text("📈 Diagrama de Pareto - Inventario Obsoleto por Plan Type (80%)", 
                       size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                chart_image
            ], spacing=15)
        else:
            self.chart_container.content = ft.Text(
                "❌ Error generando gráfico",
                size=16, color=COLORS['error']
            )
        
        self.page.update()
    
    async def _update_new_items_status(self):
        """Actualizar status de nuevos items"""
        if not self.analyzer.metrics:
            return
        
        m = self.analyzer.metrics
        new_count = m.get('new_items_count', 0)
        pending_count = m.get('pending_analysis_count', 0)
        avg_days = m.get('avg_days_in_system', 0)
        
        status_content = ft.Column([
            ft.Text("🔍 Status de Items Nuevos", size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
            ft.Row([
                ft.Container(
                    content=ft.Column([
                        ft.Text("🆕 Nuevos Items", size=14, color=COLORS['text_secondary']),
                        ft.Text(f"{new_count}", size=24, weight=ft.FontWeight.BOLD, 
                               color=COLORS['warning'] if new_count > 0 else COLORS['success'])
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=15,
                    bgcolor=COLORS['primary'],
                    border_radius=8,
                    width=150
                ),
                ft.Container(
                    content=ft.Column([
                        ft.Text("⏳ Pendientes", size=14, color=COLORS['text_secondary']),
                        ft.Text(f"{pending_count}", size=24, weight=ft.FontWeight.BOLD,
                               color=COLORS['error'] if pending_count > 0 else COLORS['success'])
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=15,
                    bgcolor=COLORS['primary'],
                    border_radius=8,
                    width=150
                ),
                ft.Container(
                    content=ft.Column([
                        ft.Text("📅 Días Promedio", size=14, color=COLORS['text_secondary']),
                        ft.Text(f"{avg_days:.1f}", size=24, weight=ft.FontWeight.BOLD, color=COLORS['accent'])
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=15,
                    bgcolor=COLORS['primary'],
                    border_radius=8,
                    width=150
                )
            ], spacing=15)
        ], spacing=15)
        
        self.new_items_status.content = status_content
        self.page.update()
    
    async def _update_new_items_table(self):
        """Actualizar tabla de items nuevos"""
        if self.analyzer.new_items is None or len(self.analyzer.new_items) == 0:
            self.table_container.visible = False
            self.page.update()
            return
        
        # Crear filas para items nuevos
        rows = []
        for _, item in self.analyzer.new_items.iterrows():
            # Truncar descripción si es muy larga
            description = str(item.get('Description', ''))
            if len(description) > 30:
                description = description[:27] + "..."
            
            # Formatear fecha
            date_added = str(item.get('date_added', ''))
            if len(date_added) > 10:
                date_added = date_added[:10]  # Solo fecha, sin hora
            
            # Crear campo de entrada para causa raíz
            root_cause_field = ft.TextField(
                value=str(item.get('root_cause', '')),
                hint_text="Ingrese causa raíz...",
                width=200,
                height=40,
                text_size=11,
                bgcolor=COLORS['surface'],
                border_color=COLORS['accent']
            )
            
            # Botón para guardar
            save_button = ft.ElevatedButton(
                text="💾",
                tooltip="Guardar causa raíz",
                on_click=lambda e, item_id=item.get('id'), field=root_cause_field: 
                    self._save_root_cause(item_id, field),
                style=ft.ButtonStyle(
                    bgcolor=COLORS['success'],
                    color=ft.Colors.WHITE,
                    padding=ft.padding.all(5)
                ),
                width=50,
                height=35
            )
            
            row = ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(str(item.get('ItemNo', '')), size=11, color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(description, size=11, color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(str(item.get('PlanType', '')), size=11, color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(f"${float(item.get('ExtOH', 0)):,.2f}", size=11, color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(date_added, size=11, color=COLORS['text_primary'])),
                    ft.DataCell(root_cause_field),
                    ft.DataCell(save_button)
                ]
            )
            rows.append(row)
        
        self.new_items_table.rows = rows
        self.table_container.visible = True
        self.page.update()
    
    async def _save_root_cause(self, item_id, root_cause_field):
        """Guardar causa raíz de un item"""
        try:
            root_cause = root_cause_field.value.strip()
            if not root_cause:
                await self._show_snackbar("❌ Por favor ingrese una causa raíz")
                return
            
            # Mostrar modal para notas adicionales
            notes_field = ft.TextField(
                hint_text="Notas adicionales del ingeniero (opcional)...",
                multiline=True,
                max_lines=3,
                width=400
            )
            
            def save_with_notes(e):
                result = self.analyzer.update_item_root_cause(
                    item_id, root_cause, notes_field.value.strip()
                )
                dialog.open = False
                self.page.update()
                
                if result["success"]:
                    asyncio.create_task(self._update_management_ui("✅ Causa raíz guardada exitosamente"))
                else:
                    asyncio.create_task(self._show_snackbar(f"❌ {result['message']}"))
            
            def cancel_save(e):
                dialog.open = False
                self.page.update()
            
            dialog = ft.AlertDialog(
                modal=True,
                title=ft.Text("💾 Confirmar Causa Raíz", color=COLORS['text_primary']),
                content=ft.Column([
                    ft.Text(f"Item: {item_id}", color=COLORS['text_secondary']),
                    ft.Text(f"Causa Raíz: {root_cause}", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                    ft.Divider(),
                    notes_field
                ], height=200),
                actions=[
                    ft.TextButton("Cancelar", on_click=cancel_save),
                    ft.ElevatedButton(
                        "💾 Guardar",
                        on_click=save_with_notes,
                        style=ft.ButtonStyle(bgcolor=COLORS['success'])
                    )
                ],
                bgcolor=COLORS['secondary']
            )
            
            self.page.dialog = dialog
            dialog.open = True
            self.page.update()
            
        except Exception as ex:
            await self._show_snackbar(f"❌ Error: {str(ex)}")



















def main():
    """Función principal de la aplicación"""
    app = ObsoleteAnalyzerApp()
    ft.app(target=app.main)

if __name__ == "__main__":
    try:
        print("🚀 Iniciando Obsolete Slow Motion Items Analyzer...")
        print("📊 Versión: 3.0 - Refactored")
        print("🔧 Desarrollado con Flet + Pandas + Matplotlib")
        print("=" * 50)
        
        # Verificar dependencias críticas
        required_modules = ['pandas', 'numpy', 'matplotlib', 'flet', 'openpyxl']
        missing_modules = []
        
        for module in required_modules:
            try:
                __import__(module)
            except ImportError:
                missing_modules.append(module)
        
        if missing_modules:
            print(f"❌ Módulos faltantes: {', '.join(missing_modules)}")
            print("💡 Instala con: pip install pandas numpy matplotlib flet openpyxl")
            exit(1)
        
        print("✅ Todas las dependencias están disponibles")
        
        # Verificar si existe R4Database
        db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        if os.path.exists(db_path):
            print(f"✅ R4Database encontrada: {db_path}")
        else:
            print(f"⚠️ R4Database no encontrada en: {db_path}")
            print("🔄 Se usarán datos de ejemplo en su lugar")
        
        print("🎯 Iniciando interfaz gráfica...")
        main()
        
    except Exception as e:
        print(f"❌ Error crítico: {e}")
        import traceback
        traceback.print_exc()
        input("Presiona Enter para salir...")

# ====================================
# CONFIGURACIONES ADICIONALES
# ====================================

# Configuración de matplotlib para mejor rendimiento
matplotlib.rcParams['figure.max_open_warning'] = 50
matplotlib.rcParams['font.size'] = 10
matplotlib.rcParams['axes.labelsize'] = 12
matplotlib.rcParams['axes.titlesize'] = 14
matplotlib.rcParams['xtick.labelsize'] = 10
matplotlib.rcParams['ytick.labelsize'] = 10
matplotlib.rcParams['legend.fontsize'] = 11

# Configuración de pandas para mejor visualización
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 50)

print("📋 Obsolete Slow Motion Items Analyzer - Configuración completa")
print("🎨 Tema: Dark Professional")
print("📊 Funcionalidades:")
print("   • Análisis de valor obsoleto por PlanType")
print("   • Gestión de items de lento movimiento")
print("   • Validación automática de items nuevos")
print("   • Interface para causa raíz por ingeniero")
print("   • Exportación de reportes completos")
print("   • Soporte para R4Database y datos de ejemplo")