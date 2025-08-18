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
        """Procesar datos y calcular métricas - VERSIÓN MEJORADA"""
        try:
            if self.slowmotion_data is None or len(self.slowmotion_data) == 0:
                print("⚠️ No hay datos de slowmotion para procesar")
                return
            
            print("🧹 Limpiando datos...")
            
            # Limpiar y convertir tipos de datos ANTES de hacer operaciones
            # PlanType
            self.slowmotion_data['PlanType'] = self.slowmotion_data['PlanType'].fillna('Unknown')
            
            # ExtOH - Limpiar y convertir a numérico
            print("🔧 Convirtiendo ExtOH a numérico...")
            self.slowmotion_data['ExtOH'] = pd.to_numeric(
                self.slowmotion_data['ExtOH'].astype(str).replace('', '0'), 
                errors='coerce'
            ).fillna(0)
            
            # QtyOH - Limpiar y convertir a numérico
            print("🔧 Convirtiendo QtyOH a numérico...")
            self.slowmotion_data['QtyOH'] = pd.to_numeric(
                self.slowmotion_data['QtyOH'].astype(str).replace('', '0'), 
                errors='coerce'
            ).fillna(0)
            
            # Cost - Limpiar y convertir a numérico
            if 'Cost' in self.slowmotion_data.columns:
                print("🔧 Convirtiendo Cost a numérico...")
                self.slowmotion_data['Cost'] = pd.to_numeric(
                    self.slowmotion_data['Cost'].astype(str).replace('', '0'), 
                    errors='coerce'
                ).fillna(0)
            
            print("✅ Tipos de datos después de limpieza:")
            print(f"   • ExtOH: {self.slowmotion_data['ExtOH'].dtype}")
            print(f"   • QtyOH: {self.slowmotion_data['QtyOH'].dtype}")
            
            # Procesar obsolete_items_data si existe
            if self.obsolete_items_data is not None:
                print("🧹 Limpiando datos de items obsoletos...")
                
                self.obsolete_items_data['PlanType'] = self.obsolete_items_data['PlanType'].fillna('Unknown')
                
                # Convertir ExtOH
                self.obsolete_items_data['ExtOH'] = pd.to_numeric(
                    self.obsolete_items_data['ExtOH'].astype(str).replace('', '0'), 
                    errors='coerce'
                ).fillna(0)
                
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
        """Actualizar tabla con nueva lógica: ItemNo + Lot + QtyOH_SUM - VERSIÓN CORREGIDA"""
        try:
            print("🔄 Actualizando tabla de items obsoletos...")
            conn = sqlite3.connect(db_path)
            
            if self.slowmotion_data is None:
                print("❌ No hay datos de slowmotion para actualizar")
                return {"success": False, "message": "No hay datos de slowmotion"}
            
            print(f"📊 Slowmotion original: {len(self.slowmotion_data)} registros")
            
            # LIMPIAR Y CONVERTIR DATOS NUMÉRICOS ANTES DEL GROUPBY
            print("🧹 Limpiando y convirtiendo tipos de datos...")
            
            # Crear una copia para no modificar los datos originales
            df_clean = self.slowmotion_data.copy()
            
            # Convertir columnas numéricas problemáticas
            numeric_columns = ['Cost', 'QtyOH', 'ExtOH', 'LT']
            
            for col in numeric_columns:
                if col in df_clean.columns:
                    print(f"🔧 Convirtiendo columna {col}...")
                    # Convertir a string primero, limpiar y luego a numeric
                    df_clean[col] = pd.to_numeric(df_clean[col].astype(str).replace('', '0'), errors='coerce').fillna(0)
            
            # Verificar tipos después de la conversión
            print("✅ Tipos de datos después de limpieza:")
            for col in numeric_columns:
                if col in df_clean.columns:
                    print(f"   • {col}: {df_clean[col].dtype}")
            
            # NUEVA LÓGICA: AGRUPAR POR ItemNo + Lot y SUMAR QtyOH
            print("🔧 Agrupando por ItemNo + Lot y sumando QtyOH...")
            
            try:
                slowmotion_grouped = df_clean.groupby(['ItemNo', 'Lot']).agg({
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
                    'Cost': 'mean',  # Promedio de costos (ahora numérico)
                    'QtyOH': 'sum',  # SUMAR CANTIDAD (ahora numérico)
                    'ExtOH': 'sum'   # SUMAR VALOR EXTENDIDO (ahora numérico)
                }).reset_index()
                
                print(f"📊 Después de agrupar: {len(slowmotion_grouped)} registros únicos")
                
            except Exception as group_error:
                print(f"❌ Error específico en groupby: {group_error}")
                
                # MÉTODO ALTERNATIVO MÁS SEGURO
                print("🔄 Intentando método alternativo de agrupación...")
                
                # Agrupar solo por las columnas clave y usar funciones más simples
                slowmotion_grouped = df_clean.groupby(['ItemNo', 'Lot']).agg({
                    'Description': 'first',
                    'UM': 'first',
                    'S': 'first', 
                    'WH': 'first',
                    'LastDate': 'first',
                    'Bin': 'first',  # Solo tomar el primero en lugar de concatenar
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
                    'Cost': 'first',    # Tomar el primer costo en lugar de promedio
                    'QtyOH': 'sum',     # SUMAR CANTIDAD
                    'ExtOH': 'sum'      # SUMAR VALOR EXTENDIDO
                }).reset_index()
                
                print(f"📊 Método alternativo exitoso: {len(slowmotion_grouped)} registros")
            
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
                init_data['department'] = ''
                init_data['responsible_person'] = ''
                init_data['assigned_date'] = ''
                
                # INSERTAR EN LOTES para evitar problemas de memoria
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
                existing_items['ItemNo'].astype(str) + '|' + 
                existing_items['Lot'].astype(str) + '|' + 
                existing_items['QtyOH'].astype(str)
            )
            
            slowmotion_keys = (
                slowmotion_grouped['ItemNo'].astype(str) + '|' + 
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
                new_items['department'] = ''
                new_items['responsible_person'] = ''
                new_items['assigned_date'] = ''
                
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
        """Crear gráfico de distribución por departamentos - VERSIÓN CORREGIDA"""
        try:
            # Verificar que tenemos datos
            if self.obsolete_items_data is None or len(self.obsolete_items_data) == 0:
                print("⚠️ No hay datos obsoletos para generar gráfico de departamentos")
                return None
            
            print(f"📊 Generando gráfico de departamentos con {len(self.obsolete_items_data)} registros")
            
            import matplotlib
            matplotlib.use('Agg')
            
            plt.style.use('dark_background')
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
            fig.patch.set_facecolor(COLORS['surface'])
            
            # GRÁFICO 1: Distribución por Departamento
            print("🔧 Procesando datos por departamento...")
            
            # Limpiar campo department - reemplazar valores vacíos/nulos
            department_clean = self.obsolete_items_data['department'].fillna('Sin Asignar')
            department_clean = department_clean.replace('', 'Sin Asignar')
            department_clean = department_clean.replace(' ', 'Sin Asignar')
            
            # Limpiar ExtOH
            extoh_clean = pd.to_numeric(self.obsolete_items_data['ExtOH'], errors='coerce').fillna(0)
            
            # Crear DataFrame temporal limpio
            dept_temp_df = pd.DataFrame({
                'department': department_clean,
                'ExtOH': extoh_clean,
                'ItemNo': self.obsolete_items_data['ItemNo']
            })
            
            # Agrupar por departamento
            dept_data = dept_temp_df.groupby('department').agg({
                'ExtOH': 'sum',
                'ItemNo': 'count'
            }).reset_index()
            dept_data.columns = ['Department', 'Total_Value', 'Item_Count']
            
            print(f"📈 Departamentos encontrados: {len(dept_data)}")
            print(f"📋 Departamentos: {dept_data['Department'].tolist()}")
            
            if len(dept_data) > 0 and dept_data['Total_Value'].sum() > 0:
                colors = ['#4ECDC4', '#45B7D1', '#FF6B6B', '#FFA726', '#AB47BC', '#26C6DA', '#66BB6A', '#EC407A']
                
                # Pie chart de departamentos
                wedges, texts, autotexts = ax1.pie(
                    dept_data['Total_Value'].values, 
                    labels=dept_data['Department'].values,
                    colors=colors[:len(dept_data)], 
                    autopct='%1.1f%%',
                    startangle=90,
                    textprops={'color': 'white', 'fontsize': 10}
                )
                
                ax1.set_title('Distribución de Valor por Departamento', 
                            color='white', fontsize=16, fontweight='bold', pad=20)
                
                print("✅ Gráfico 1 (pie chart) generado")
            else:
                ax1.text(0.5, 0.5, 'No hay datos de departamentos\ncon valores válidos', 
                        ha='center', va='center', color='white', fontsize=14,
                        transform=ax1.transAxes)
                ax1.set_title('Sin Datos de Departamentos', color='white', fontsize=16)
            
            # GRÁFICO 2: Status por Departamento
            print("🔧 Procesando datos por status...")
            
            try:
                # Limpiar campo status
                status_clean = self.obsolete_items_data['status'].fillna('Sin Status')
                status_clean = status_clean.replace('', 'Sin Status')
                
                # Crear DataFrame para status
                status_temp_df = pd.DataFrame({
                    'department': department_clean,
                    'status': status_clean
                })
                
                # Crear tabla cruzada
                status_dept = status_temp_df.groupby(['department', 'status']).size().unstack(fill_value=0)
                
                print(f"📊 Status cruzado generado: {status_dept.shape}")
                
                if not status_dept.empty and status_dept.values.sum() > 0:
                    # Crear gráfico de barras
                    status_colors = ['#FF6B6B', '#FFA726', '#4ECDC4', '#45B7D1', '#AB47BC', '#26C6DA']
                    
                    status_dept.plot(
                        kind='bar', 
                        ax=ax2, 
                        color=status_colors[:len(status_dept.columns)],
                        alpha=0.8
                    )
                    
                    ax2.set_title('Status por Departamento', color='white', fontsize=16, fontweight='bold')
                    ax2.set_xlabel('Departamento', color='white', fontsize=14)
                    ax2.set_ylabel('Cantidad de Items', color='white', fontsize=14)
                    ax2.tick_params(colors='white')
                    ax2.legend(loc='upper right', fancybox=True, shadow=True)
                    
                    # Rotar etiquetas del eje X
                    ax2.tick_params(axis='x', rotation=45)
                    
                    print("✅ Gráfico 2 (barras) generado")
                else:
                    ax2.text(0.5, 0.5, 'No hay datos de status\npor departamento', 
                            ha='center', va='center', color='white', fontsize=14,
                            transform=ax2.transAxes)
                    ax2.set_title('Sin Datos de Status', color='white', fontsize=16)
                    
            except Exception as status_error:
                print(f"⚠️ Error en gráfico de status: {status_error}")
                ax2.text(0.5, 0.5, f'Error procesando status:\n{str(status_error)}', 
                        ha='center', va='center', color='yellow', fontsize=12,
                        transform=ax2.transAxes)
                ax2.set_title('Error en Status', color='yellow', fontsize=16)
            
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
            
            print("✅ Gráfico de departamentos generado exitosamente")
            return chart_base64
            
        except Exception as e:
            print(f"❌ Error creando gráfico departamentos: {e}")
            import traceback
            traceback.print_exc()
            return None

# MÉTODO ALTERNATIVO MÁS SIMPLE SI EL ANTERIOR FALLA
    def create_simple_department_chart(self):
        """Crear gráfico simple de departamentos como fallback"""
        try:
            if self.obsolete_items_data is None or len(self.obsolete_items_data) == 0:
                return None
            
            import matplotlib
            matplotlib.use('Agg')
            
            plt.style.use('dark_background')
            fig, ax = plt.subplots(1, 1, figsize=(12, 8))
            fig.patch.set_facecolor(COLORS['surface'])
            
            # Análisis simple por departamento
            dept_counts = self.obsolete_items_data['department'].fillna('Sin Asignar').value_counts()
            
            if len(dept_counts) > 0:
                # Gráfico de barras simple
                bars = ax.bar(range(len(dept_counts)), dept_counts.values, 
                            color=['#4ECDC4', '#45B7D1', '#FF6B6B', '#FFA726', '#AB47BC'][:len(dept_counts)])
                
                ax.set_xticks(range(len(dept_counts)))
                ax.set_xticklabels(dept_counts.index, rotation=45, ha='right')
                ax.set_title('Cantidad de Items por Departamento', color='white', fontsize=16, fontweight='bold')
                ax.set_ylabel('Cantidad de Items', color='white', fontsize=14)
                ax.tick_params(colors='white')
                
                # Agregar valores en las barras
                for bar, value in zip(bars, dept_counts.values):
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height + max(dept_counts.values)*0.01,
                        f'{value}', ha='center', va='bottom', color='white', fontsize=12)
            else:
                ax.text(0.5, 0.5, 'No hay datos para mostrar', ha='center', va='center',
                    color='white', fontsize=16, transform=ax.transAxes)
            
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
            print(f"❌ Error en gráfico simple: {e}")
            return None

    # MÉTODO PARA DEBUGGEAR LOS DATOS DE DEPARTAMENTOS
    def debug_department_data(self):
        """Debuggear datos de departamentos"""
        try:
            if self.obsolete_items_data is None:
                print("❌ No hay datos obsoletos")
                return
            
            print(f"📊 Total registros: {len(self.obsolete_items_data)}")
            print(f"📋 Columnas disponibles: {list(self.obsolete_items_data.columns)}")
            
            # Verificar columna department
            if 'department' in self.obsolete_items_data.columns:
                dept_info = self.obsolete_items_data['department'].value_counts(dropna=False)
                print(f"🏢 Distribución de departamentos:")
                print(dept_info)
                
                # Verificar valores nulos/vacíos
                null_count = self.obsolete_items_data['department'].isnull().sum()
                empty_count = (self.obsolete_items_data['department'] == '').sum()
                print(f"❓ Valores nulos: {null_count}")
                print(f"❓ Valores vacíos: {empty_count}")
            else:
                print("❌ Columna 'department' no existe")
            
            # Verificar columna status
            if 'status' in self.obsolete_items_data.columns:
                status_info = self.obsolete_items_data['status'].value_counts(dropna=False)
                print(f"📊 Distribución de status:")
                print(status_info)
            else:
                print("❌ Columna 'status' no existe")
            
            # Verificar ExtOH
            if 'ExtOH' in self.obsolete_items_data.columns:
                extoh_sum = pd.to_numeric(self.obsolete_items_data['ExtOH'], errors='coerce').sum()
                print(f"💰 Suma total ExtOH: ${extoh_sum:,.2f}")
            else:
                print("❌ Columna 'ExtOH' no existe")
                
        except Exception as e:
            print(f"❌ Error en debug: {e}")
            import traceback
            traceback.print_exc()


    def update_item_assignment(self, item_id, department, responsible_person, 
                            db_path=r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"):
        """Actualizar asignación de departamento y responsable"""
        try:
            print(f"🔧 DB Update - ID: {item_id}, Dept: {department}, Resp: {responsible_person}")
            
            conn = sqlite3.connect(db_path)
            
            update_query = """
            UPDATE newObsolet_slowmotion_items 
            SET department = ?, responsible_person = ?, assigned_date = ?, status = 'Assigned'
            WHERE id = ?
            """
            
            assigned_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            cursor = conn.cursor()
            cursor.execute(update_query, (department, responsible_person, assigned_date, item_id))
            
            rows_affected = cursor.rowcount
            conn.commit()
            conn.close()
            
            print(f"✅ DB Update exitoso - Filas afectadas: {rows_affected}")
            
            if rows_affected > 0:
                return {"success": True, "message": f"Asignación actualizada (ID: {item_id})"}
            else:
                return {"success": False, "message": f"No se encontró item con ID: {item_id}"}
            
        except Exception as e:
            error_msg = f"Error actualizando DB: {str(e)}"
            print(f"❌ {error_msg}")
            return {"success": False, "message": error_msg}

    def _verify_save_assignment_flow(self):
        """Método de verificación para debuggear el flujo completo"""
        try:
            print("🔍 VERIFICANDO FLUJO DE GUARDADO:")
            
            # 1. Verificar que el analyzer tiene el método
            has_method = hasattr(self.analyzer, 'update_item_assignment')
            print(f"1. ✅ Analyzer tiene update_item_assignment: {has_method}")
            
            # 2. Verificar que _save_assignment no es async
            import inspect
            is_async = inspect.iscoroutinefunction(self._save_assignment)
            print(f"2. ✅ _save_assignment es síncrono: {not is_async}")
            
            # 3. Verificar datos disponibles
            has_data = self.analyzer.obsolete_items_data is not None
            print(f"3. ✅ Hay datos obsoletos: {has_data}")
            
            if has_data:
                data_count = len(self.analyzer.obsolete_items_data)
                print(f"   📊 Registros disponibles: {data_count}")
            
            # 4. Verificar helper methods
            has_success = hasattr(self, '_show_success_safely')
            has_error = hasattr(self, '_show_error_safely')
            print(f"4. ✅ Helper methods: success={has_success}, error={has_error}")
            
            # 5. Verificar tabla sync method
            has_sync_update = hasattr(self, '_update_department_table_sync')
            print(f"5. ✅ Método sync update: {has_sync_update}")
            
            print("🎯 FLUJO DE GUARDADO VERIFICADO")
            
        except Exception as e:
            print(f"❌ Error en verificación: {e}")

    def _test_save_assignment(self, e):
        """Probar guardado con datos de prueba"""
        try:
            print("🧪 INICIANDO PRUEBA DE GUARDADO...")
            
            # Verificar flujo primero
            self._verify_save_assignment_flow()
            
            # Crear controles dummy para prueba
            dummy_dept = ft.Dropdown(value="VENTAS")
            dummy_resp = ft.TextField(value="Usuario Prueba")
            
            # Probar con ID válido (usar el primer registro disponible)
            if self.analyzer.obsolete_items_data is not None and len(self.analyzer.obsolete_items_data) > 0:
                test_id = self.analyzer.obsolete_items_data.iloc[0]['id']
                print(f"🧪 Probando con ID: {test_id}")
                
                self._save_assignment(test_id, dummy_dept, dummy_resp)
            else:
                print("❌ No hay datos para probar")
                self._show_error_safely("❌ No hay datos cargados para probar")
                
        except Exception as ex:
            print(f"❌ Error en prueba: {ex}")

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

    def _save_assignment(self, item_id, dept_dropdown, responsible_field):
        """Guardar asignación de departamento y responsable - VERSIÓN SÍNCRONA CORREGIDA"""
        try:
            department = dept_dropdown.value if dept_dropdown.value else ""
            responsible = responsible_field.value.strip() if responsible_field.value else ""
            
            print(f"💾 Guardando asignación - ID: {item_id}, Dept: {department}, Resp: {responsible}")
            
            # Llamar al método del analyzer (que ya es síncrono)
            result = self.analyzer.update_item_assignment(item_id, department, responsible)
            
            if result["success"]:
                # Mostrar mensaje de éxito
                self._show_success_safely(f"✅ Asignación guardada: {department} - {responsible}")
                
                # Recargar datos
                self.analyzer.load_data_from_db()
                print(f"✅ Datos recargados después de guardar ID {item_id}")
                
            else:
                # Mostrar error
                self._show_error_safely(f"❌ {result['message']}")
                
        except Exception as ex:
            error_msg = f"❌ Error guardando asignación: {str(ex)}"
            print(error_msg)
            self._show_error_safely(error_msg)

    async def _export_by_department(self, e):
        """Exportar por departamento"""
        await self._show_snackbar("🚧 Función de exportación por departamento en desarrollo")

    async def _filter_by_department(self, e):
        """Filtrar por departamento"""
        await self._show_snackbar("🚧 Filtro por departamento en desarrollo")

    async def _filter_by_status(self, e):
        """Filtrar por status"""
        await self._show_snackbar("🚧 Filtro por status en desarrollo")

    def import_massive_update(self, file_path, db_path=r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"):
        """Importar y aplicar update masivo desde Excel"""
        try:
            print(f"📥 Importando update masivo desde: {file_path}")
            
            # Leer archivo Excel
            try:
                import_df = pd.read_excel(file_path, sheet_name='Items_Para_Editar')
                print(f"📊 Registros leídos del Excel: {len(import_df)}")
            except Exception as read_error:
                return {"success": False, "message": f"Error leyendo Excel: {str(read_error)}"}
            
            # Validaciones básicas
            if 'id' not in import_df.columns:
                return {"success": False, "message": "❌ Falta columna 'id' requerida en el archivo"}
            
            # Filtrar solo registros con ID válido
            valid_updates = import_df[
                (import_df['id'].notna()) & 
                (import_df['id'] != '') &
                (import_df['id'] > 0)
            ].copy()
            
            print(f"📊 Registros válidos para update: {len(valid_updates)}")
            
            if len(valid_updates) == 0:
                return {"success": False, "message": "❌ No hay registros válidos para actualizar"}
            
            # Conectar a BD y hacer updates
            conn = sqlite3.connect(db_path)
            updated_count = 0
            errors = []
            
            print("🔄 Iniciando updates en base de datos...")
            
            for index, row in valid_updates.iterrows():
                try:
                    item_id = int(row['id'])
                    
                    # Construir UPDATE query dinámicamente
                    update_fields = []
                    update_values = []
                    
                    # Campos editables que pueden ser actualizados
                    editable_fields = {
                        'root_cause': 'root_cause',
                        'engineer_notes': 'engineer_notes', 
                        'status': 'status',
                        'department': 'department',
                        'responsible_person': 'responsible_person'
                    }
                    
                    for excel_col, db_col in editable_fields.items():
                        if excel_col in row and pd.notna(row[excel_col]):
                            value = str(row[excel_col]).strip()
                            if value != '':
                                update_fields.append(f"{db_col} = ?")
                                update_values.append(value)
                    
                    if update_fields:
                        # Agregar timestamp de update
                        update_fields.append("assigned_date = ?")
                        update_values.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                        
                        # Si se actualiza root_cause o status, marcar como no nuevo
                        if any('root_cause' in field or 'status' in field for field in update_fields):
                            update_fields.append("is_new = ?")
                            update_values.append(0)
                        
                        update_query = f"""
                        UPDATE newObsolet_slowmotion_items 
                        SET {', '.join(update_fields)}
                        WHERE id = ?
                        """
                        update_values.append(item_id)
                        
                        cursor = conn.cursor()
                        cursor.execute(update_query, update_values)
                        
                        if cursor.rowcount > 0:
                            updated_count += 1
                            print(f"✅ Actualizado ID {item_id}")
                        else:
                            errors.append(f"ID {item_id} no encontrado en base de datos")
                    
                except Exception as row_error:
                    error_msg = f"Error en fila {index + 2} (ID {row.get('id', 'N/A')}): {str(row_error)}"
                    errors.append(error_msg)
                    print(f"❌ {error_msg}")
            
            conn.commit()
            conn.close()
            
            # Recargar datos en memoria
            print("🔄 Recargando datos...")
            self.load_data_from_db(db_path)
            
            result_message = f"✅ Update masivo completado: {updated_count}/{len(valid_updates)} registros actualizados"
            if errors:
                result_message += f"\n⚠️ {len(errors)} errores encontrados"
                print("❌ Errores encontrados:")
                for error in errors[:5]:  # Mostrar solo los primeros 5
                    print(f"   • {error}")
            
            return {
                "success": True,
                "message": result_message,
                "updated_count": updated_count,
                "total_records": len(valid_updates),
                "errors": errors[:10]  # Máximo 10 errores en el resultado
            }
            
        except Exception as e:
            error_msg = f"Error general en import masivo: {str(e)}"
            print(f"❌ {error_msg}")
            import traceback
            traceback.print_exc()
            return {"success": False, "message": error_msg}

    def export_for_massive_update(self):
        """Exportar datos para edición masiva en Excel"""
        try:
            print("🔄 Iniciando exportación para update masivo...")
            
            if self.obsolete_items_data is None or len(self.obsolete_items_data) == 0:
                return {"success": False, "message": "No hay datos para exportar"}
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            export_path = f"Massive_Update_Template_{timestamp}.xlsx"
            
            print(f"📁 Archivo de exportación: {export_path}")
            
            # Preparar datos para exportación
            export_data = self.obsolete_items_data.copy()
            
            # Limpiar y convertir ExtOH para ordenamiento
            export_data['ExtOH'] = pd.to_numeric(export_data['ExtOH'], errors='coerce').fillna(0)
            
            # Ordenar por ExtOH descendente (valores más altos primero)
            export_data = export_data.sort_values('ExtOH', ascending=False)
            
            print(f"📊 Datos ordenados por ExtOH: {len(export_data)} registros")
            
            # Seleccionar columnas para exportación
            # Columnas requeridas y opcionales
            required_columns = ['id']  # CRÍTICO: ID es necesario para el update
            display_columns = [
                'ItemNo', 'Description', 'PlanType', 'ExtOH', 'QtyOH', 
                'Lot', 'Bin', 'Entity', 'WH', 'LastDate'
            ]
            editable_columns = [
                'status', 'root_cause', 'engineer_notes', 
                'department', 'responsible_person'
            ]
            info_columns = ['date_added', 'is_new']
            
            # Construir lista de columnas disponibles
            available_columns = required_columns.copy()
            
            for col_list in [display_columns, editable_columns, info_columns]:
                for col in col_list:
                    if col in export_data.columns and col not in available_columns:
                        available_columns.append(col)
            
            # Crear DataFrame final
            export_df = export_data[available_columns].copy()
            
            print(f"📋 Columnas incluidas: {len(available_columns)}")
            print(f"📋 Columnas: {available_columns}")
            
            # Formatear datos para mejor visualización
            if 'ExtOH' in export_df.columns:
                export_df['ExtOH'] = export_df['ExtOH'].round(2)
            
            if 'QtyOH' in export_df.columns:
                export_df['QtyOH'] = pd.to_numeric(export_df['QtyOH'], errors='coerce').fillna(0)
            
            if 'date_added' in export_df.columns:
                export_df['date_added'] = pd.to_datetime(export_df['date_added'], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # Rellenar valores vacíos
            export_df = export_df.fillna('')
            
            # Crear archivo Excel con múltiples hojas
            print("📝 Creando archivo Excel...")
            
            try:
                with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                    # Hoja 1: Datos principales para editar
                    export_df.to_excel(writer, sheet_name='Items_Para_Editar', index=False)
                    print("✅ Hoja 'Items_Para_Editar' creada")
                    
                    # Hoja 2: Instrucciones detalladas
                    instructions_data = {
                        'INSTRUCCIONES PARA UPDATE MASIVO': [
                            '🔧 PASOS PARA UPDATE MASIVO:',
                            '',
                            '1. 📝 EDITAR SOLO ESTAS COLUMNAS:',
                            '   • root_cause (Causa raíz del problema)',
                            '   • engineer_notes (Notas del ingeniero)',
                            '   • status (Estado: Pending Analysis, Analyzed, Assigned, Resolved)',
                            '   • department (PRODUCCION, ALMACEN, CALIDAD, INGENIERIA, COMPRAS, VENTAS, MANTENIMIENTO)',
                            '   • responsible_person (Nombre del responsable)',
                            '',
                            '2. ❌ NO MODIFICAR:',
                            '   • La columna "id" (CRÍTICA para identificar registros)',
                            '   • No agregar filas nuevas',
                            '   • No eliminar filas existentes',
                            '   • No modificar ItemNo, Description, ExtOH, etc.',
                            '',
                            '3. 💾 GUARDAR Y CARGAR:',
                            '   • Guarda este archivo Excel modificado',
                            '   • Ve a la aplicación → "📤 3. Importar Cambios"',
                            '   • Selecciona este archivo modificado',
                            '',
                            '4. ✅ VALIDACIONES:',
                            '   • El sistema validará los datos antes del update',
                            '   • Se mostrarán errores si hay problemas',
                            '   • Solo se actualizarán registros válidos',
                            '',
                            '5. 📊 DATOS ACTUALES:',
                            f'   • Total de registros: {len(export_df)}',
                            f'   • Ordenados por ExtOH (mayor a menor)',
                            f'   • Fecha de exportación: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}',
                            '',
                            '⚠️ IMPORTANTE: Mantén el formato Excel original'
                        ]
                    }
                    
                    instructions_df = pd.DataFrame(instructions_data)
                    instructions_df.to_excel(writer, sheet_name='INSTRUCCIONES', index=False)
                    print("✅ Hoja 'INSTRUCCIONES' creada")
                    
                    # Hoja 3: Valores válidos y ejemplos
                    valid_values_data = {
                        'Status_Validos': [
                            'Pending Analysis',
                            'Analyzed', 
                            'Assigned',
                            'Resolved',
                            '(vacío = no cambiar)',
                            ''
                        ],
                        'Departamentos_Validos': [
                            'PRODUCCION',
                            'ALMACEN',
                            'CALIDAD', 
                            'INGENIERIA',
                            'COMPRAS',
                            'VENTAS'
                        ],
                        'Ejemplos_Causa_Raiz': [
                            'Obsolescencia tecnológica',
                            'Cambio de diseño',
                            'Fin de ciclo de producto',
                            'Error de pronóstico',
                            'Cambio de proveedor',
                            'Sobrante de proyecto'
                        ],
                        'Ejemplos_Notas_Ingeniero': [
                            'Validado con ingeniería',
                            'Requiere disposición especial',
                            'Evaluar uso alternativo',
                            'Contactar cliente final',
                            'Verificar inventario físico',
                            'Programar reunión'
                        ]
                    }
                    
                    # Hacer que todas las listas tengan la misma longitud
                    max_len = max(len(v) for v in valid_values_data.values())
                    for key in valid_values_data:
                        while len(valid_values_data[key]) < max_len:
                            valid_values_data[key].append('')
                    
                    valid_values_df = pd.DataFrame(valid_values_data)
                    valid_values_df.to_excel(writer, sheet_name='Valores_Validos', index=False)
                    print("✅ Hoja 'Valores_Validos' creada")
                    
                    # Hoja 4: Resumen de datos
                    summary_data = {
                        'Métrica': [
                            'Total de registros',
                            'Registros con status "Pending Analysis"',
                            'Registros con causa raíz vacía',
                            'Valor total ExtOH',
                            'Promedio ExtOH por item',
                            'Top 10% valor ExtOH',
                            'Items nuevos (is_new = 1)',
                            'Departamentos únicos asignados'
                        ],
                        'Valor': [
                            len(export_df),
                            len(export_df[export_df.get('status', '') == 'Pending Analysis']),
                            len(export_df[export_df.get('root_cause', '') == '']),
                            f"${export_df.get('ExtOH', 0).sum():,.2f}",
                            f"${export_df.get('ExtOH', 0).mean():,.2f}",
                            f"${export_df.get('ExtOH', 0).quantile(0.9):,.2f}",
                            len(export_df[export_df.get('is_new', 0) == 1]),
                            export_df.get('department', '').nunique() if 'department' in export_df.columns else 0
                        ]
                    }
                    
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='Resumen_Datos', index=False)
                    print("✅ Hoja 'Resumen_Datos' creada")
                    
            except Exception as excel_error:
                print(f"❌ Error creando Excel: {excel_error}")
                return {"success": False, "message": f"Error creando archivo Excel: {str(excel_error)}"}
            
            print(f"✅ Archivo Excel creado exitosamente: {export_path}")
            
            # Abrir carpeta donde se guardó el archivo
            try:
                full_path = os.path.abspath(export_path)
                print(f"📂 Ruta completa: {full_path}")
                
                if platform.system() == "Windows":
                    subprocess.Popen(f'explorer /select,"{full_path}"')
                    print("📁 Carpeta abierta en Windows Explorer")
                elif platform.system() == "Darwin":  # macOS
                    subprocess.Popen(["open", "-R", full_path])
                    print("📁 Carpeta abierta en Finder")
                else:  # Linux
                    subprocess.Popen(["xdg-open", os.path.dirname(full_path)])
                    print("📁 Carpeta abierta en explorador")
                    
            except Exception as open_error:
                print(f"⚠️ No se pudo abrir carpeta: {open_error}")
            
            return {
                "success": True, 
                "message": f"Archivo exportado: {export_path}",
                "file_path": full_path,
                "records_count": len(export_df)
            }
            
        except Exception as e:
            error_msg = f"Error exportando para update masivo: {str(e)}"
            print(f"❌ {error_msg}")
            import traceback
            traceback.print_exc()
            return {"success": False, "message": error_msg}

# =========================================================


class ObsoleteAnalyzerApp:
    def __init__(self):
        """Inicializar aplicación con todas las variables necesarias"""
        self.analyzer = ObsoleteSlowMotionAnalyzer()
        self.page = None
        self.loading_modal = None
        
        # Variables de paginación
        self.current_page = 1
        self.items_per_page = 25
        self.filtered_data = None
        
        # Referencias a componentes UI
        self.search_field = None
        self.new_items_table = None
        self.table_container = None
        self.pagination_controls = None
        self.new_items_status = None
        self.metrics_container = None
        self.chart_container = None
        
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
        """Crear tab de gestión de items obsoletos - VERSIÓN CORREGIDA"""
        
        # Panel de control principal
        management_control = ft.Container(
            content=ft.Column([
                ft.Text("🔧 Gestión de Items Obsoletos", size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                
                # Fila 1: Botones principales
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
                        text="📊 Exportar Reporte Completo",
                        on_click=self._export_obsolete_report,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['accent'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=280
                    )
                ], spacing=15),
                
                # Sección Update Masivo
                ft.Divider(color=COLORS['accent']),
                ft.Text("📤 Update Masivo en Excel", size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Text("1. Exporta → 2. Edita en Excel → 3. Importa", size=12, color=COLORS['text_secondary']),
                
                ft.Row([
                    ft.ElevatedButton(
                        text="📥 1. Exportar para Edición",
                        on_click=self._export_for_massive_edit,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['warning'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=280,
                        tooltip="Descargar Excel para editar masivamente"
                    ),
                    ft.ElevatedButton(
                        text="📤 3. Importar Cambios",
                        on_click=self._import_massive_update,
                        style=ft.ButtonStyle(
                            bgcolor=COLORS['success'],
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=20, vertical=15)
                        ),
                        width=280,
                        tooltip="Subir Excel modificado para aplicar cambios"
                    )
                ], spacing=15),
                
                # Botones de debug/prueba
                
                
                # Instrucciones
                ft.Container(
                    content=ft.Column([
                        ft.Text("📋 Instrucciones:", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                        ft.Text("• Exporta el archivo Excel", size=10, color=COLORS['text_secondary']),
                        ft.Text("• Edita: root_cause, engineer_notes, status, department, responsible_person", size=10, color=COLORS['text_secondary']),
                        ft.Text("• NO modifiques la columna 'id'", size=10, color=COLORS['text_secondary']),
                        ft.Text("• Guarda e importa los cambios", size=10, color=COLORS['text_secondary'])
                    ], spacing=2),
                    padding=10,
                    bgcolor=COLORS['primary'],
                    border_radius=8,
                    margin=ft.margin.only(top=10)
                )
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
        
        # Tabla simple
        self.new_items_table = ft.DataTable(
            columns=[
                ft.DataColumn(label=ft.Text("Item No", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Descripción", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("PlanType", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Valor ExtOH", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Status", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Causa Raíz", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary'])),
                ft.DataColumn(label=ft.Text("Acciones", size=12, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']))
            ],
            rows=[],
            bgcolor=COLORS['secondary'],
            border_radius=12
        )
        
        self.table_container = ft.Container(
            content=ft.Column([
                ft.Text("📋 Items Obsoletos (Ordenados por Valor ExtOH)", size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Container(
                    content=ft.Column([
                        self.new_items_table
                    ], scroll=ft.ScrollMode.AUTO),
                    bgcolor=COLORS['secondary'],
                    border_radius=12,
                    padding=10,
                    height=500
                )
            ]),
            visible=False
        )
        
        # Inicializar variables
        self.current_page = 1
        self.items_per_page = 25
        self.filtered_data = None
        
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
        """Actualizar tabla de items obsoletos - VERSIÓN MEJORADA"""
        await self._show_loading()
        
        try:
            print("🔄 Iniciando actualización de items obsoletos...")
            result = self.analyzer.update_obsolete_items_table()
            
            if result["success"]:
                await self._update_management_ui(f"✅ {result['message']}")
                
                # Mostrar mensaje más informativo
                new_count = result['new_items_count']
                if new_count > 0:
                    await self._show_snackbar(f"✅ {result['message']} - {new_count} nuevos items encontrados")
                else:
                    await self._show_snackbar(f"✅ {result['message']}")
            else:
                await self._show_snackbar(f"❌ {result['message']}")
                
        except Exception as ex:
            print(f"❌ Error en actualización: {str(ex)}")
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
        """Actualizar UI del tab de gestión - VERSIÓN CORREGIDA"""
        if self.analyzer.obsolete_items_data is not None:
            # Actualizar status de nuevos items
            await self._update_new_items_status()
            
            # Actualizar tabla de items con todos los datos ordenados
            await self._update_new_items_table()
            
            await self._show_snackbar(status_message)

    async def _import_massive_update_simple(self, e):
        """Método alternativo más simple para importar"""
        try:
            await self._show_snackbar("📁 Abriendo selector de archivo...")
            
            # Crear file picker simple
            file_picker = ft.FilePicker(
                on_result=lambda e: self._handle_file_selection_simple(e)
            )
            
            self.page.overlay.clear()
            self.page.overlay.append(file_picker)
            self.page.update()
            
            # Abrir file picker
            file_picker.pick_files(
                dialog_title="Seleccionar Excel de Update Masivo",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx", "xls"]
            )
            
        except Exception as ex:
            await self._show_snackbar(f"❌ Error: {str(ex)}")

    async def _process_massive_update(self, file_path):
        """Procesar el archivo de update masivo seleccionado"""
        await self._show_loading()
        
        try:
            print(f"🔄 Procesando archivo: {file_path}")
            
            # Verificar que el archivo existe
            if not os.path.exists(file_path):
                await self._show_snackbar("❌ Archivo no encontrado")
                return
            
            # Llamar al método del analyzer
            result = self.analyzer.import_massive_update(file_path)
            
            if result["success"]:
                success_msg = f"✅ Update masivo completado!\n"
                success_msg += f"📊 {result['updated_count']} de {result['total_records']} registros actualizados"
                
                await self._show_snackbar(success_msg)
                
                # Recargar la UI
                await self._update_management_ui("Update masivo completado exitosamente")
                
                # Mostrar errores si los hay
                if 'errors' in result and result['errors']:
                    error_msg = f"⚠️ Se encontraron {len(result['errors'])} errores:\n"
                    for error in result['errors'][:3]:  # Mostrar solo los primeros 3
                        error_msg += f"• {error}\n"
                    print(error_msg)
            else:
                await self._show_snackbar(f"❌ Error en update masivo: {result['message']}")
                
        except Exception as ex:
            await self._show_snackbar(f"❌ Error procesando archivo: {str(ex)}")
            print(f"Error en _process_massive_update: {ex}")
            import traceback
            traceback.print_exc()
        
        await self._hide_loading()

    async def _export_for_massive_edit(self, e):
        """Exportar datos para edición masiva - VERSIÓN DEBUG"""
        print("🔍 DEBUG: Botón de exportar clickeado")
        
        try:
            # Verificar que tenemos datos
            if self.analyzer.obsolete_items_data is None:
                print("❌ DEBUG: No hay datos obsoletos cargados")
                await self._show_snackbar("❌ No hay datos cargados. Primero actualiza los items obsoletos.")
                return
            
            if len(self.analyzer.obsolete_items_data) == 0:
                print("❌ DEBUG: Datos obsoletos están vacíos")
                await self._show_snackbar("❌ No hay datos para exportar")
                return
            
            print(f"✅ DEBUG: Tenemos {len(self.analyzer.obsolete_items_data)} registros para exportar")
            
            await self._show_loading()
            
            # Verificar que el método export existe
            if not hasattr(self.analyzer, 'export_for_massive_update'):
                print("❌ DEBUG: Método export_for_massive_update no existe")
                await self._show_snackbar("❌ Error: Método de exportación no encontrado")
                await self._hide_loading()
                return
            
            print("✅ DEBUG: Llamando a export_for_massive_update...")
            result = self.analyzer.export_for_massive_update()
            print(f"📊 DEBUG: Resultado: {result}")
            
            if result["success"]:
                success_msg = f"✅ {result['message']}\n📊 {result['records_count']} registros exportados"
                await self._show_snackbar(success_msg)
                print("✅ DEBUG: Exportación exitosa")
            else:
                await self._show_snackbar(f"❌ {result['message']}")
                print(f"❌ DEBUG: Error en exportación: {result['message']}")
                
        except Exception as ex:
            error_msg = f"❌ Error: {str(ex)}"
            await self._show_snackbar(error_msg)
            print(f"❌ DEBUG: Excepción: {ex}")
            import traceback
            traceback.print_exc()
        
        finally:
            await self._hide_loading()

    async def _export_simple_test(self, e):
        """Método simple para probar exportación"""
        print("🧪 PRUEBA: Iniciando exportación simple")
        
        try:
            if self.analyzer.obsolete_items_data is None or len(self.analyzer.obsolete_items_data) == 0:
                await self._show_snackbar("❌ No hay datos para exportar")
                return
            
            # Crear archivo Excel simple
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Simple_Export_Test_{timestamp}.xlsx"
            
            # Usar primeros 50 registros
            test_data = self.analyzer.obsolete_items_data.head(50).copy()
            
            # Columnas básicas que sabemos que existen
            if 'ItemNo' in test_data.columns:
                # Crear DataFrame simple
                simple_export = test_data[['ItemNo']].copy()
                if 'Description' in test_data.columns:
                    simple_export['Description'] = test_data['Description']
                if 'ExtOH' in test_data.columns:
                    simple_export['ExtOH'] = test_data['ExtOH']
                
                # Exportar
                simple_export.to_excel(filename, index=False)
                await self._show_snackbar(f"✅ Prueba exitosa: {filename} creado")
                print(f"✅ PRUEBA: {filename} creado con {len(simple_export)} registros")
            else:
                await self._show_snackbar("❌ No se encontró columna ItemNo")
                
        except Exception as ex:
            await self._show_snackbar(f"❌ Error en prueba: {str(ex)}")
            print(f"❌ PRUEBA: {ex}")

    async def _show_debug_info(self, e):
        """Mostrar información de debug"""
        try:
            debug_lines = []
            debug_lines.append(f"Analyzer existe: {self.analyzer is not None}")
            
            if self.analyzer:
                debug_lines.append(f"Datos obsoletos: {self.analyzer.obsolete_items_data is not None}")
                if self.analyzer.obsolete_items_data is not None:
                    debug_lines.append(f"Registros: {len(self.analyzer.obsolete_items_data)}")
                    debug_lines.append(f"Columnas: {len(self.analyzer.obsolete_items_data.columns)}")
                debug_lines.append(f"Método export: {hasattr(self.analyzer, 'export_for_massive_update')}")
            
            debug_msg = " | ".join(debug_lines)
            await self._show_snackbar(f"🔍 {debug_msg}")
            
            # Print completo en consola
            print("🔍 DEBUG INFO COMPLETA:")
            for line in debug_lines:
                print(f"   • {line}")
            
            if self.analyzer and self.analyzer.obsolete_items_data is not None:
                print(f"   • Columnas disponibles: {list(self.analyzer.obsolete_items_data.columns)}")
            
        except Exception as ex:
            await self._show_snackbar(f"❌ Error debug: {str(ex)}")

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
                color="#ff9800",  # Naranja693
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
    
    async def _update_new_items_table(self):
        """Actualizar tabla de items nuevos - VERSIÓN CORREGIDA"""
        if self.analyzer.obsolete_items_data is None or len(self.analyzer.obsolete_items_data) == 0:
            if hasattr(self, 'table_container') and self.table_container:
                self.table_container.visible = False
                self.page.update()
            return
        
        # Verificar que los componentes UI existen
        if not hasattr(self, 'new_items_table') or self.new_items_table is None:
            print("⚠️ Componentes de tabla no inicializados, saltando actualización")
            return
        
        # Usar todos los datos ordenados por ExtOH descendente
        try:
            # Limpiar y convertir ExtOH a numérico si es necesario
            df_clean = self.analyzer.obsolete_items_data.copy()
            df_clean['ExtOH'] = pd.to_numeric(df_clean['ExtOH'], errors='coerce').fillna(0)
            
            # Ordenar por ExtOH descendente
            all_data = df_clean.sort_values('ExtOH', ascending=False)
            
            self.filtered_data = all_data
            self.current_page = 1
            
            await self._update_table_with_data(all_data)
            
        except Exception as e:
            print(f"❌ Error actualizando tabla: {e}")
            # Fallback: usar método simple sin paginación
            await self._update_table_simple()

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

    def _create_department_tab(self):
        """Crear tab de gestión por departamentos - CON BÚSQUEDA"""
        
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
        
        # NUEVO: Panel de búsqueda y filtros
        self.department_search_field = ft.TextField(
            hint_text="🔍 Buscar por Item No, Descripción, Lote...",
            width=300,
            height=45,
            text_size=14,
            bgcolor=COLORS['surface'],
            border_color=COLORS['accent'],
            on_change=self._on_department_search_change,
            on_submit=self._perform_department_search
        )
        
        search_and_filters = ft.Container(
            content=ft.Column([
                ft.Text("🔍 Búsqueda y Filtros", size=16, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Row([
                    # Campo de búsqueda
                    self.department_search_field,
                    
                    # Botón de búsqueda
                    ft.ElevatedButton(
                        text="🔍 Buscar",
                        on_click=self._perform_department_search,
                        style=ft.ButtonStyle(bgcolor=COLORS['accent']),
                        height=45
                    ),
                    
                    # Botón limpiar búsqueda
                    ft.ElevatedButton(
                        text="🧹 Limpiar",
                        on_click=self._clear_department_search,
                        style=ft.ButtonStyle(bgcolor=COLORS['warning']),
                        height=45
                    ),
                    
                    # Dropdown de departamentos
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
                    
                    # Dropdown de status
                    ft.Dropdown(
                        label="Filtrar por Status",
                        options=[
                            ft.dropdown.Option("", "Todos"),
                            ft.dropdown.Option("Pending Analysis", "Pendiente Análisis"),
                            ft.dropdown.Option("Analyzed", "Analizado"),
                            ft.dropdown.Option("Assigned", "Asignado"),
                            ft.dropdown.Option("Historical", "Histórico")
                        ],
                        width=200,
                        on_change=self._filter_by_status
                    )
                ], spacing=15, wrap=True)
            ], spacing=10),
            padding=20,
            bgcolor=COLORS['secondary'],
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Información de resultados
        self.department_results_info = ft.Container(
            content=ft.Text("📋 Carga datos para ver resultados", 
                        size=14, color=COLORS['text_secondary']),
            padding=10,
            bgcolor=COLORS['primary'],
            border_radius=8,
            margin=ft.margin.only(bottom=10)
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
                self.department_results_info,
                ft.Container(
                    content=ft.Column([
                        self.department_table
                    ], scroll=ft.ScrollMode.AUTO),
                    bgcolor=COLORS['secondary'],
                    border_radius=12,
                    padding=10,
                    height=500
                )
            ]),
            visible=False
        )
        
        # Variables para filtros y búsqueda
        self.department_filtered_data = None
        self.department_search_term = ""
        
        return ft.Column([
            department_control,
            self.department_chart_container,
            search_and_filters,
            self.department_table_container
        ], scroll=ft.ScrollMode.AUTO)

    def _on_department_search_change(self, e):
        """Manejar cambios en el campo de búsqueda (búsqueda en tiempo real)"""
        # Búsqueda automática después de 1 segundo de inactividad
        pass

    # PASO 1: Convertir _perform_department_search a completamente síncrono
    def _perform_department_search(self, e):
        """Realizar búsqueda en la tabla de departamentos - VERSIÓN SÍNCRONA"""
        try:
            if not hasattr(self, 'department_search_field') or not self.department_search_field:
                print("❌ Campo de búsqueda no inicializado")
                return
                
            search_term = self.department_search_field.value.strip().lower()
            self.department_search_term = search_term
            
            print(f"🔍 Búsqueda departamental: '{search_term}'")
            
            if self.analyzer.obsolete_items_data is None or len(self.analyzer.obsolete_items_data) == 0:
                self._show_error_safely("❌ No hay datos para buscar")
                return
            
            # Filtrar datos
            if search_term:
                # Buscar en múltiples campos
                filtered_data = self.analyzer.obsolete_items_data[
                    (self.analyzer.obsolete_items_data['ItemNo'].astype(str).str.lower().str.contains(search_term, na=False)) |
                    (self.analyzer.obsolete_items_data['Description'].astype(str).str.lower().str.contains(search_term, na=False)) |
                    (self.analyzer.obsolete_items_data['Lot'].astype(str).str.lower().str.contains(search_term, na=False)) |
                    (self.analyzer.obsolete_items_data['Bin'].astype(str).str.lower().str.contains(search_term, na=False))
                ]
                
                self.department_filtered_data = filtered_data
                result_count = len(filtered_data)
                
                # Actualizar info de resultados
                if hasattr(self, 'department_results_info'):
                    self.department_results_info.content = ft.Text(
                        f"🔍 Búsqueda: '{search_term}' • {result_count} resultados encontrados",
                        size=14, color=COLORS['accent']
                    )
                
                self._show_success_safely(f"🔍 {result_count} resultados para '{search_term}'")
                
            else:
                # Sin filtro, mostrar todos
                self.department_filtered_data = None
                if hasattr(self, 'department_results_info'):
                    self.department_results_info.content = ft.Text(
                        f"📋 Mostrando todos los registros • {len(self.analyzer.obsolete_items_data)} total",
                        size=14, color=COLORS['text_secondary']
                    )
            
            # Actualizar tabla de forma SÍNCRONA
            self._update_department_table_sync()
            
        except Exception as ex:
            print(f"❌ Error en búsqueda departamental: {ex}")
            self._show_error_safely(f"❌ Error en búsqueda: {str(ex)}")

    # PASO 3: Corregir _clear_department_search para que sea completamente síncrono
    def _clear_department_search(self, e):
        """Limpiar búsqueda y mostrar todos los registros - VERSIÓN SÍNCRONA"""
        try:
            print("🧹 Limpiando búsqueda departamental...")
            
            if hasattr(self, 'department_search_field'):
                self.department_search_field.value = ""
            
            self.department_search_term = ""
            self.department_filtered_data = None
            
            # Actualizar info
            if hasattr(self, 'department_results_info') and self.analyzer.obsolete_items_data is not None:
                self.department_results_info.content = ft.Text(
                    f"📋 Mostrando todos los registros • {len(self.analyzer.obsolete_items_data)} total",
                    size=14, color=COLORS['text_secondary']
                )
            
            # Actualizar tabla de forma SÍNCRONA
            self._update_department_table_sync()
            
            self._show_success_safely("✅ Búsqueda limpiada")
            
        except Exception as ex:
            print(f"❌ Error limpiando búsqueda: {ex}")
            self._show_error_safely(f"❌ Error limpiando búsqueda: {str(ex)}")

    def _update_department_table_sync(self):
        """Actualizar tabla de departamentos de forma SÍNCRONA"""
        try:
            print("🔄 Actualizando tabla departamental (síncrono)...")
            
            # Determinar qué datos usar
            if hasattr(self, 'department_filtered_data') and self.department_filtered_data is not None:
                data_to_show = self.department_filtered_data
                print(f"📊 Usando datos filtrados: {len(data_to_show)} registros")
            elif self.analyzer.obsolete_items_data is not None:
                data_to_show = self.analyzer.obsolete_items_data
                print(f"📊 Usando todos los datos: {len(data_to_show)} registros")
            else:
                print("❌ No hay datos para mostrar")
                if hasattr(self, 'department_table_container'):
                    self.department_table_container.visible = False
                    self.page.update()
                return
            
            if len(data_to_show) == 0:
                print("⚠️ No hay resultados para mostrar")
                if hasattr(self, 'department_table_container'):
                    self.department_table_container.visible = False
                self._show_error_safely("❌ No se encontraron resultados")
                self.page.update()
                return
            
            # Crear filas (limitar a 100 para performance)
            rows = []
            display_data = data_to_show.head(100)  # Mostrar máximo 100 registros
            print(f"📋 Creando {len(display_data)} filas para la tabla...")
            
            for _, item in display_data.iterrows():
                
                # Dropdown para departamento
                dept_dropdown = ft.Dropdown(
                    value=str(item.get('department', '')),
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
                item_id = item.get('id')
                save_btn = ft.ElevatedButton(
                    text="💾",
                    tooltip="Guardar asignación",
                    on_click=lambda e, id=item_id, dept=dept_dropdown, resp=responsible_field: 
                        self._save_assignment(id, dept, resp),
                    style=ft.ButtonStyle(bgcolor=COLORS['success']),
                    width=40,
                    height=30
                )
                
                row = ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(str(item.get('ItemNo', ''))[:15], size=10, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(str(item.get('Description', ''))[:25], size=10, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(f"${float(item.get('ExtOH', 0)):,.0f}", size=10, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(str(item.get('status', ''))[:12], size=10, color=COLORS['text_primary'])),
                        ft.DataCell(dept_dropdown),
                        ft.DataCell(responsible_field),
                        ft.DataCell(save_btn)
                    ]
                )
                rows.append(row)
            
            # Actualizar tabla
            if hasattr(self, 'department_table'):
                self.department_table.rows = rows
                print(f"✅ Tabla actualizada con {len(rows)} filas")
                
            if hasattr(self, 'department_table_container'):
                self.department_table_container.visible = True
                
            # Actualizar info de resultados
            if hasattr(self, 'department_results_info'):
                total_filtered = len(data_to_show)
                showing = len(display_data)
                
                if hasattr(self, 'department_search_term') and self.department_search_term:
                    info_text = f"🔍 '{self.department_search_term}' • Mostrando {showing} de {total_filtered} resultados"
                else:
                    info_text = f"📋 Mostrando {showing} de {total_filtered} registros"
                    
                self.department_results_info.content = ft.Text(info_text, size=14, color=COLORS['accent'])
                print(f"📊 Info actualizada: {info_text}")
            
            # Actualizar la página
            self.page.update()
            print(f"✅ Tabla departamental actualizada exitosamente: {len(rows)} filas mostradas")
            
        except Exception as ex:
            print(f"❌ Error actualizando tabla departamental: {ex}")
            import traceback
            traceback.print_exc()
            self._show_error_safely(f"❌ Error actualizando tabla: {str(ex)}")

    async def _update_department_table_with_search(self):
        """Actualizar tabla de departamentos considerando filtros de búsqueda"""
        try:
            # Determinar qué datos usar
            if self.department_filtered_data is not None:
                data_to_show = self.department_filtered_data
            elif self.analyzer.obsolete_items_data is not None:
                data_to_show = self.analyzer.obsolete_items_data
            else:
                if hasattr(self, 'department_table_container'):
                    self.department_table_container.visible = False
                return
            
            if len(data_to_show) == 0:
                # No hay resultados
                if hasattr(self, 'department_table_container'):
                    self.department_table_container.visible = False
                self._show_error_safely("❌ No se encontraron resultados")
                return
            
            # Crear filas (limitar a 100 para performance)
            rows = []
            display_data = data_to_show.head(100)  # Mostrar máximo 100 registros
            
            for _, item in display_data.iterrows():
                
                # Dropdown para departamento
                dept_dropdown = ft.Dropdown(
                    value=str(item.get('department', '')),
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
                
                # Botón de guardar - CORREGIDO
                item_id = item.get('id')
                save_btn = ft.ElevatedButton(
                    text="💾",
                    tooltip="Guardar asignación",
                    on_click=lambda e, id=item_id, dept=dept_dropdown, resp=responsible_field: 
                        self._save_assignment(id, dept, resp),  # LLAMADA SÍNCRONA
                    style=ft.ButtonStyle(bgcolor=COLORS['success']),
                    width=40,
                    height=30
                )
                
                row = ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(str(item.get('ItemNo', ''))[:15], size=10, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(str(item.get('Description', ''))[:25], size=10, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(f"${float(item.get('ExtOH', 0)):,.0f}", size=10, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(str(item.get('status', ''))[:12], size=10, color=COLORS['text_primary'])),
                        ft.DataCell(dept_dropdown),
                        ft.DataCell(responsible_field),
                        ft.DataCell(save_btn)
                    ]
                )
                rows.append(row)
            
            # Actualizar tabla
            if hasattr(self, 'department_table'):
                self.department_table.rows = rows
                
            if hasattr(self, 'department_table_container'):
                self.department_table_container.visible = True
                
            # Actualizar info de resultados
            if hasattr(self, 'department_results_info'):
                total_filtered = len(data_to_show)
                showing = len(display_data)
                
                if self.department_search_term:
                    info_text = f"🔍 '{self.department_search_term}' • Mostrando {showing} de {total_filtered} resultados"
                else:
                    info_text = f"📋 Mostrando {showing} de {total_filtered} registros"
                    
                self.department_results_info.content = ft.Text(info_text, size=14, color=COLORS['accent'])
            
            self.page.update()
            print(f"✅ Tabla departamental actualizada: {len(rows)} filas mostradas")
            
        except Exception as ex:
            print(f"❌ Error actualizando tabla departamental: {ex}")
            import traceback
            traceback.print_exc()

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
        """Actualizar tabla de departamentos - LLAMA A VERSIÓN SÍNCRONA"""
        print("🔄 _update_department_table llamado (async -> sync)")
        self._update_department_table_sync()

    def _test_department_search(self, e):
        """Método de prueba para la búsqueda departamental"""
        try:
            # Simular búsqueda
            if hasattr(self, 'department_search_field'):
                self.department_search_field.value = "025"
                self._perform_department_search(None)
            else:
                print("❌ Campo de búsqueda no encontrado")
                
        except Exception as ex:
            print(f"❌ Error en prueba de búsqueda: {ex}")

    # PASO 6: Botón de prueba para agregar temporalmente
    def _create_test_search_button(self):
        """Crear botón de prueba para búsqueda (temporal)"""
        return ft.ElevatedButton(
            text="🧪 Probar Búsqueda",
            on_click=self._test_department_search,
            style=ft.ButtonStyle(bgcolor=COLORS['warning']),
            width=200
        )

    # PASO 7: Método para inicializar variables de búsqueda
    def _initialize_department_search_variables(self):
        """Inicializar variables de búsqueda departamental"""
        self.department_filtered_data = None
        self.department_search_term = ""
        print("✅ Variables de búsqueda departamental inicializadas")

    # PASO 8: Versión mejorada del método de búsqueda con mejor manejo de errores
    def _perform_department_search_improved(self, e):

        """Versión mejorada de búsqueda departamental"""
        try:
            print("🔍 Iniciando búsqueda departamental mejorada...")
            
            # Verificar prerrequisitos
            if not hasattr(self, 'department_search_field'):
                print("❌ Campo de búsqueda no inicializado")
                self._show_error_safely("❌ Campo de búsqueda no disponible")
                return
                
            if not self.department_search_field:
                print("❌ Referencia al campo de búsqueda es None")
                return
                
            if self.analyzer.obsolete_items_data is None:
                print("❌ No hay datos obsoletos cargados")
                self._show_error_safely("❌ No hay datos cargados. Actualiza primero los items obsoletos.")
                return
                
            if len(self.analyzer.obsolete_items_data) == 0:
                print("❌ Datos obsoletos están vacíos")
                self._show_error_safely("❌ No hay datos para buscar")
                return
            
            # Obtener término de búsqueda
            search_term = self.department_search_field.value.strip().lower()
            self.department_search_term = search_term
            
            print(f"🔍 Término de búsqueda: '{search_term}'")
            print(f"📊 Total de registros disponibles: {len(self.analyzer.obsolete_items_data)}")
            
            # Realizar filtrado
            if search_term:
                try:
                    # Crear máscaras de búsqueda para cada campo
                    mask_itemno = self.analyzer.obsolete_items_data['ItemNo'].astype(str).str.lower().str.contains(search_term, na=False)
                    mask_desc = self.analyzer.obsolete_items_data['Description'].astype(str).str.lower().str.contains(search_term, na=False)
                    mask_lot = self.analyzer.obsolete_items_data['Lot'].astype(str).str.lower().str.contains(search_term, na=False)
                    mask_bin = self.analyzer.obsolete_items_data['Bin'].astype(str).str.lower().str.contains(search_term, na=False)
                    
                    # Combinar máscaras
                    combined_mask = mask_itemno | mask_desc | mask_lot | mask_bin
                    
                    # Aplicar filtro
                    filtered_data = self.analyzer.obsolete_items_data[combined_mask]
                    
                    self.department_filtered_data = filtered_data
                    result_count = len(filtered_data)
                    
                    print(f"✅ Filtrado completado: {result_count} resultados")
                    
                    # Actualizar info de resultados
                    if hasattr(self, 'department_results_info'):
                        self.department_results_info.content = ft.Text(
                            f"🔍 Búsqueda: '{search_term}' • {result_count} resultados encontrados",
                            size=14, color=COLORS['accent']
                        )
                    
                    if result_count > 0:
                        self._show_success_safely(f"🔍 {result_count} resultados para '{search_term}'")
                    else:
                        self._show_error_safely(f"❌ No se encontraron resultados para '{search_term}'")
                    
                except Exception as filter_error:
                    print(f"❌ Error en filtrado: {filter_error}")
                    self._show_error_safely(f"❌ Error filtrando datos: {str(filter_error)}")
                    return
                    
            else:
                # Sin filtro, mostrar todos
                self.department_filtered_data = None
                if hasattr(self, 'department_results_info'):
                    self.department_results_info.content = ft.Text(
                        f"📋 Mostrando todos los registros • {len(self.analyzer.obsolete_items_data)} total",
                        size=14, color=COLORS['text_secondary']
                    )
                print("📋 Mostrando todos los registros")
            
            # Actualizar tabla
            self._update_department_table_sync()
            
        except Exception as ex:
            error_msg = f"❌ Error general en búsqueda: {str(ex)}"
            print(error_msg)
            import traceback
            traceback.print_exc()
            self._show_error_safely(error_msg)

    def _save_assignment(self, item_id, dept_dropdown, responsible_field):
        """Guardar asignación de departamento y responsable - VERSIÓN FINAL SÍNCRONA"""
        try:
            # Obtener valores de los controles
            department = dept_dropdown.value if dept_dropdown.value else ""
            responsible = responsible_field.value.strip() if responsible_field.value else ""
            
            print(f"💾 GUARDANDO - ID: {item_id}, Dept: '{department}', Resp: '{responsible}'")
            
            # Validaciones básicas
            if not item_id:
                print("❌ ID de item inválido")
                self._show_error_safely("❌ ID de item inválido")
                return
            
            # Llamar al método del analyzer (que es síncrono)
            result = self.analyzer.update_item_assignment(item_id, department, responsible)
            
            if result["success"]:
                success_msg = f"✅ Guardado: {department} - {responsible}"
                print(success_msg)
                
                # Mostrar mensaje de éxito
                self._show_success_safely(success_msg)
                
                # Recargar datos en el analyzer
                print("🔄 Recargando datos...")
                self.analyzer.load_data_from_db()
                
                # Actualizar la tabla para reflejar cambios
                print("🔄 Actualizando tabla...")
                self._update_department_table_sync()
                
                print("✅ Guardado completado exitosamente")
                
            else:
                error_msg = f"❌ Error DB: {result['message']}"
                print(error_msg)
                self._show_error_safely(error_msg)
                
        except Exception as ex:
            error_msg = f"❌ Error guardando: {str(ex)}"
            print(error_msg)
            import traceback
            traceback.print_exc()
            self._show_error_safely(error_msg)

    async def _export_by_department(self, e):
        """Exportar por departamento"""
        await self._show_snackbar("🚧 Función de exportación por departamento en desarrollo")

    async def _filter_by_department(self, e):
        """Filtrar por departamento"""
        await self._show_snackbar("🚧 Filtro por departamento en desarrollo")

    async def _filter_by_status(self, e):
        """Filtrar por status"""
        await self._show_snackbar("🚧 Filtro por status en desarrollo")

    async def _import_massive_update(self, e):
        """Importar update masivo - VERSIÓN CORREGIDA"""
        try:
            await self._show_snackbar("📁 Abriendo selector de archivo...")
            
            # Variable para almacenar el resultado
            self._selected_file_path = None
            
            def on_file_picker_result(e: ft.FilePickerResultEvent):
                """Callback del FilePicker - NO ASÍNCRONO"""
                if e.files and len(e.files) > 0:
                    selected_path = e.files[0].path
                    print(f"📁 Archivo seleccionado: {selected_path}")
                    
                    # Almacenar la ruta para procesarla después
                    self._selected_file_path = selected_path
                    
                    # Programar el procesamiento en el bucle principal
                    # Usar page.add para trigger un update que ejecute el procesamiento
                    self.page.add(ft.Container(visible=False))  # Dummy container para trigger update
                    self.page.update()
                    
                    # Llamar al procesamiento de forma síncrona usando threading
                    import threading
                    
                    def process_in_thread():
                        """Procesar en un hilo separado y notificar al bucle principal"""
                        try:
                            # Crear un nuevo bucle de eventos para este hilo
                            import asyncio
                            loop = asyncio.new_event_loop()
                            asyncio.set_event_loop(loop)
                            
                            # Ejecutar el procesamiento asíncrono
                            loop.run_until_complete(self._process_massive_update_sync(selected_path))
                            
                            # Cerrar el bucle
                            loop.close()
                            
                        except Exception as ex:
                            print(f"❌ Error en hilo de procesamiento: {ex}")
                            # Mostrar error de forma segura
                            self._show_error_safely(f"Error procesando archivo: {str(ex)}")
                    
                    # Iniciar el hilo de procesamiento
                    thread = threading.Thread(target=process_in_thread, daemon=True)
                    thread.start()
                else:
                    print("❌ No se seleccionó ningún archivo")
            
            # Crear file picker
            file_picker = ft.FilePicker(on_result=on_file_picker_result)
            
            # Limpiar overlay previo y agregar el nuevo
            self.page.overlay.clear()
            self.page.overlay.append(file_picker)
            self.page.update()
            
            # Abrir file picker
            file_picker.pick_files(
                dialog_title="Seleccionar archivo Excel de Update Masivo",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx", "xls"]
            )
            
        except Exception as ex:
            await self._show_snackbar(f"❌ Error abriendo selector: {str(ex)}")
            print(f"❌ Error en _import_massive_update: {ex}")

    async def _process_massive_update_sync(self, file_path):
        """Versión síncrona del procesamiento para usar en hilos"""
        try:
            print(f"🔄 Procesando archivo: {file_path}")
            
            # Verificar que el archivo existe
            if not os.path.exists(file_path):
                self._show_error_safely("❌ Archivo no encontrado")
                return
            
            # Llamar al método del analyzer (no asíncrono)
            result = self.analyzer.import_massive_update(file_path)
            
            if result["success"]:
                success_msg = f"✅ Update masivo completado!\n"
                success_msg += f"📊 {result['updated_count']} de {result['total_records']} registros actualizados"
                
                self._show_success_safely(success_msg)
                
                # Recargar datos en el analyzer
                self.analyzer.load_data_from_db()
                
                # Programar actualización de UI
                self._schedule_ui_update()
                
                # Mostrar errores si los hay
                if 'errors' in result and result['errors']:
                    error_msg = f"⚠️ Se encontraron {len(result['errors'])} errores:\n"
                    for error in result['errors'][:3]:  # Mostrar solo los primeros 3
                        error_msg += f"• {error}\n"
                    print(error_msg)
            else:
                self._show_error_safely(f"❌ Error en update masivo: {result['message']}")
                
        except Exception as ex:
            error_msg = f"❌ Error procesando archivo: {str(ex)}"
            self._show_error_safely(error_msg)
            print(f"Error en _process_massive_update_sync: {ex}")
            import traceback
            traceback.print_exc()

    def _show_success_safely(self, message):
        """Mostrar mensaje de éxito de forma thread-safe"""
        try:
            def show_success():
                try:
                    snackbar = ft.SnackBar(
                        content=ft.Text(message, color=ft.Colors.WHITE),
                        bgcolor=COLORS['success'],
                        duration=3000
                    )
                    self.page.snack_bar = snackbar
                    snackbar.open = True
                    self.page.update()
                    print(f"✅ Snackbar mostrado: {message}")
                except Exception as e:
                    print(f"Error mostrando snackbar éxito: {e}")
                    print(f"Mensaje: {message}")
            
            try:
                if hasattr(self.page, 'run_thread_safe'):
                    self.page.run_thread_safe(show_success)
                else:
                    show_success()
            except:
                print(f"✅ FALLBACK: {message}")
        except Exception as e:
            print(f"Error en _show_success_safely: {e}")
            print(f"Mensaje original: {message}")

    def _show_error_safely(self, message):
        """Mostrar mensaje de error de forma thread-safe"""
        try:
            def show_error():
                try:
                    snackbar = ft.SnackBar(
                        content=ft.Text(message, color=ft.Colors.WHITE),
                        bgcolor=COLORS['error'],
                        duration=5000
                    )
                    self.page.snack_bar = snackbar
                    snackbar.open = True
                    self.page.update()
                    print(f"❌ Snackbar mostrado: {message}")
                except Exception as e:
                    print(f"Error mostrando snackbar error: {e}")
                    print(f"Mensaje: {message}")
            
            try:
                if hasattr(self.page, 'run_thread_safe'):
                    self.page.run_thread_safe(show_error)
                else:
                    show_error()
            except:
                print(f"❌ FALLBACK: {message}")
        except Exception as e:
            print(f"Error en _show_error_safely: {e}")
            print(f"Mensaje original: {message}")

    def _schedule_ui_update(self):
        """Programar actualización de UI de forma thread-safe"""
        try:
            def update_ui():
                try:
                    # Crear una tarea asíncrona para actualizar la UI
                    import asyncio
                    
                    async def update_management():
                        await self._update_management_ui("Update masivo completado exitosamente")
                    
                    # Verificar si hay un bucle corriendo
                    try:
                        loop = asyncio.get_running_loop()
                        # Si hay bucle, crear la tarea
                        loop.create_task(update_management())
                    except RuntimeError:
                        # No hay bucle, ejecutar de forma síncrona
                        print("✅ Datos actualizados, UI se actualizará en la próxima interacción")
                        
                except Exception as e:
                    print(f"Error actualizando UI: {e}")
            
            if hasattr(self.page, 'run_thread_safe'):
                self.page.run_thread_safe(update_ui)
            else:
                print("✅ Update completado, UI se actualizará automáticamente")
                
        except Exception as e:
            print(f"Error en _schedule_ui_update: {e}")

    def _on_search_change(self, e):
        """Manejar cambios en el campo de búsqueda (búsqueda en tiempo real)"""
        # Búsqueda automática después de 1 segundo de inactividad
        pass

    async def _perform_search(self, e):
        """Realizar búsqueda"""
        search_term = self.search_field.value.strip()
        await self._show_loading()
        try:
            self.filtered_data = self.analyzer.search_items(search_term)
            self.current_page = 1
            await self._update_table_with_data(self.filtered_data)
            
            if search_term:
                await self._show_snackbar(f"🔍 Búsqueda: {len(self.filtered_data)} resultados para '{search_term}'")
            else:
                await self._show_snackbar("📋 Mostrando todos los items")
        except Exception as ex:
            await self._show_snackbar(f"❌ Error en búsqueda: {str(ex)}")
        await self._hide_loading()

    async def _show_all_items(self, e):
        """Mostrar todos los items"""
        self.search_field.value = ""
        await self._perform_search(e)

    def _previous_page(self, e):
        """Página anterior"""
        if self.current_page > 1:
            self.current_page -= 1
            asyncio.create_task(self._update_current_page())

    def _next_page(self, e):
        """Página siguiente"""
        data = self.filtered_data if self.filtered_data is not None else self.analyzer.obsolete_items_data
        if data is not None:
            total_pages = math.ceil(len(data) / self.items_per_page)
            if self.current_page < total_pages:
                self.current_page += 1
                asyncio.create_task(self._update_current_page())

    async def _update_current_page(self):
        """Actualizar página actual"""
        data = self.filtered_data if self.filtered_data is not None else self.analyzer.obsolete_items_data
        await self._update_table_with_data(data)

    async def _update_table_with_data(self, data):
        """Actualizar tabla con datos específicos y paginación - VERSIÓN CORREGIDA"""
        try:
            if data is None or len(data) == 0:
                if hasattr(self, 'table_container') and self.table_container:
                    self.table_container.visible = False
                    self.page.update()
                return
            
            # Verificar que tenemos items_per_page inicializado
            if not hasattr(self, 'items_per_page'):
                self.items_per_page = 25
            if not hasattr(self, 'current_page'):
                self.current_page = 1
            
            # Calcular paginación
            total_records = len(data)
            total_pages = math.ceil(total_records / self.items_per_page)
            start_idx = (self.current_page - 1) * self.items_per_page
            end_idx = start_idx + self.items_per_page
            
            page_data = data.iloc[start_idx:end_idx]
            
            # Crear filas
            rows = []
            for _, item in page_data.iterrows():
                # Truncar descripción
                description = str(item.get('Description', ''))
                if len(description) > 25:
                    description = description[:22] + "..."
                
                # Campo de causa raíz
                root_cause_field = ft.TextField(
                    value=str(item.get('root_cause', '')),
                    hint_text="Causa raíz...",
                    width=180,
                    height=40,
                    text_size=11,
                    bgcolor=COLORS['surface'],
                    border_color=COLORS['accent']
                )
                
                # Botón guardar
                save_button = ft.ElevatedButton(
                    text="💾",
                    tooltip="Guardar causa raíz",
                    on_click=lambda e, item_id=item.get('id'), field=root_cause_field: 
                        self._save_root_cause(item_id, field),
                    style=ft.ButtonStyle(bgcolor=COLORS['success']),
                    width=50,
                    height=35
                )
                
                row = ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(str(item.get('ItemNo', '')), size=11, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(description, size=11, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(str(item.get('PlanType', '')), size=11, color=COLORS['text_primary'])),
                        ft.DataCell(ft.Text(f"${float(item.get('ExtOH', 0)):,.0f}", size=11, weight=ft.FontWeight.BOLD, color=COLORS['accent'])),
                        ft.DataCell(ft.Text(str(item.get('status', '')), size=11, color=COLORS['text_primary'])),
                        ft.DataCell(root_cause_field),
                        ft.DataCell(save_button)
                    ]
                )
                rows.append(row)
            
            self.new_items_table.rows = rows
            
            # Actualizar controles de paginación solo si existen
            if hasattr(self, 'pagination_controls') and self.pagination_controls:
                page_info = f"Página {self.current_page} de {total_pages}"
                record_info = f"Mostrando {start_idx + 1}-{min(end_idx, total_records)} de {total_records} registros"
                
                self.pagination_controls.controls[1].value = page_info
                self.pagination_controls.controls[4].value = record_info
                
                # Habilitar/deshabilitar botones
                self.pagination_controls.controls[0].disabled = (self.current_page <= 1)
                self.pagination_controls.controls[2].disabled = (self.current_page >= total_pages)
            
            if hasattr(self, 'table_container') and self.table_container:
                self.table_container.visible = True
            
            self.page.update()
            print(f"✅ Tabla con paginación actualizada: página {self.current_page}, {len(page_data)} registros mostrados")
            
        except Exception as e:
            print(f"❌ Error en paginación, usando método simple: {e}")
            # Fallback al método simple
            await self._update_table_simple()

    async def _verify_methods(self, e):
        """Verificar que todos los métodos existen - BOTÓN DE PRUEBA"""
        try:
            messages = []
            
            # Verificar métodos en analyzer (SIN guión bajo)
            messages.append(f"export_for_massive_update: {hasattr(self.analyzer, 'export_for_massive_update')}")
            messages.append(f"import_massive_update: {hasattr(self.analyzer, 'import_massive_update')}")
            
            # Verificar métodos en app (CON guión bajo)
            messages.append(f"_export_for_massive_edit: {hasattr(self, '_export_for_massive_edit')}")
            messages.append(f"_import_massive_update: {hasattr(self, '_import_massive_update')}")
            messages.append(f"_process_massive_update: {hasattr(self, '_process_massive_update')}")
            
            verification_msg = " | ".join(messages)
            await self._show_snackbar(f"🔍 Métodos: {verification_msg}")
            
            print("🔍 VERIFICACIÓN DE MÉTODOS:")
            for msg in messages:
                print(f"   • {msg}")
                
        except Exception as ex:
            await self._show_snackbar(f"❌ Error verificando: {str(ex)}")

# =========================================================

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