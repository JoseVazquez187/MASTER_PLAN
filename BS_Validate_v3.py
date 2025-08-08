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
import matplotlib.pyplot as plt
import matplotlib
import io
import base64
import math
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

class BackScheduleAnalyzer:
    def __init__(self):
        self.df = None
        self.metrics = {}

    def load_data_from_db(self, db_path=r"C:\Users\J.Vazquez\Desktop\PARCHE EXPEDITE\BS_Analisis_2025_08_04.db"):
        """Cargar datos desde la base de datos SQLite"""
        try:
            print(f"Intentando conectar a: {db_path}")
            conn = sqlite3.connect(db_path)
            query = """
                SELECT EntityGroup, ItemNo, Description, DemandSource, AC, ShipDate, OH, MLIKCode, 
                    LT_TOTAL, ReqDate, ReqDate_Original
                FROM expedite
                """
            self.df = pd.read_sql_query(query, conn)
            conn.close()
            print(f"✅ Datos cargados exitosamente: {len(self.df)} registros")
            self._process_data()
            return True
        except Exception as e:
            print(f"❌ Error al cargar desde DB: {e}")
            print("🔄 Usando datos de ejemplo...")
            self._load_sample_data()
            return False
    
    def _load_sample_data(self):
        """Datos de ejemplo mejorados para pruebas"""
        np.random.seed(42)
        n_records = 100
        
        items = [f"02-10-{100+i:03d}-10" for i in range(20)]
        entity_groups = ['MX', 'US', 'CA', 'BR']
        demand_sources = ['SO', 'PO', 'TO', 'MO']
        
        data = {
            "EntityGroup": np.random.choice(entity_groups, n_records),
            "ItemNo": np.random.choice(items, n_records),
            "Description": [f"Descripción Item {i}" for i in range(n_records)],
            "DemandSource": np.random.choice(demand_sources, n_records),
            "AC": np.random.choice(['AC1', 'AC2', 'AC3'], n_records),
            "ShipDate": pd.date_range(start='2025-07-01', periods=n_records, freq='D'),
            "OH": np.random.randint(0, 1000, n_records),
            "MLIKCode": np.random.choice(['A', 'B', 'C'], n_records),
            "LT_TOTAL": np.random.randint(1, 30, n_records),
            "ReqDate": [],
            "ReqDate_Original": []
        }
        
        base_date = datetime(2025, 7, 1)
        for i in range(n_records):
            original = base_date + timedelta(days=i + np.random.randint(-5, 15))
            diff_days = int(np.random.choice([-7, -3, -1, 0, 1, 2, 5, 10], 
                                        p=[0.1, 0.15, 0.2, 0.3, 0.15, 0.05, 0.03, 0.02]))
            req = original + timedelta(days=diff_days)
            
            data["ReqDate"].append(req)
            data["ReqDate_Original"].append(original)
        
        data["ShipDate"] = [rd - timedelta(days=int(np.random.randint(1, 10))) 
                            for rd in data["ReqDate"]]
        
        self.df = pd.DataFrame(data)
        coincidences = np.random.choice(self.df.index, size=n_records//4, replace=False)
        self.df.loc[coincidences, "DemandSource"] = self.df.loc[coincidences, "ItemNo"]
        self._process_data()

    def _process_data(self):
        """Procesar y limpiar los datos"""
        if self.df is None:
            return
            
        print(f"📊 Procesando {len(self.df)} registros...")
        
        try:
            self.df["ReqDate"] = pd.to_datetime(self.df["ReqDate"], errors='coerce')
            self.df["ReqDate_Original"] = self.df["ReqDate_Original"].replace(['Req-Date', 'req-date', 'REQ-DATE'], pd.NaT)
            self.df["ReqDate_Original"] = pd.to_datetime(self.df["ReqDate_Original"], errors='coerce')
            self.df["ShipDate"] = pd.to_datetime(self.df["ShipDate"], errors='coerce')
            
            initial_count = len(self.df)
            self.df = self.df.dropna(subset=['ReqDate', 'ReqDate_Original', 'ShipDate'])
            final_count = len(self.df)
            
            print(f"📊 Registros válidos: {final_count}/{initial_count} ({100*final_count/initial_count:.1f}%)")
            
            if len(self.df) == 0:
                print("❌ No hay registros válidos después de limpiar fechas")
                self._load_sample_data()
                return
            
            # CORREGIDO: Ahora calculamos las diferencias correctas
            self.df["Diff_BackSchedule"] = (self.df["ReqDate"] - self.df["ReqDate_Original"]).dt.days.astype(int)
            self.df["Diff_Ship_Req"] = (self.df["ShipDate"] - self.df["ReqDate"]).dt.days.astype(int)  # ShipDate - ReqDate (BackSchedule)
            self.df["Diff_Ship_Original"] = (self.df["ShipDate"] - self.df["ReqDate_Original"]).dt.days.astype(int)  # ShipDate - ReqDate_Original
            
            # NUEVA COLUMNA: Material_Type basada en DemandSource vs ItemNo
            self.df["Material_Type"] = self.df.apply(
                lambda row: "Invoiceable" if row["DemandSource"] == row["ItemNo"] else "Raw Material", 
                axis=1
            )



            conditions = [
                (self.df["Diff_BackSchedule"] == 0),
                (self.df["Diff_BackSchedule"].abs() <= 1),
                (self.df["Diff_BackSchedule"].abs() <= 3),
                (self.df["Diff_BackSchedule"].abs() <= 7),
            ]
            choices = ['Exacto', '±1 día', '±3 días', '±7 días']
            self.df["Precision_Category"] = np.select(conditions, choices, default='> 7 días')
            
            self._calculate_metrics()
            print("✅ Procesamiento completado exitosamente")
            
        except Exception as e:
            print(f"❌ Error procesando datos reales: {e}")
            print("🔄 Usando datos de ejemplo...")
            self._load_sample_data()
    
    def _calculate_metrics(self):
        """Calcular KPIs avanzados"""
        total = len(self.df)
        if total == 0:
            return
            
        exact_match = int((self.df["Diff_BackSchedule"] == 0).sum())
        within_1_day = int((self.df["Diff_BackSchedule"].abs() <= 1).sum())
        within_3_days = int((self.df["Diff_BackSchedule"].abs() <= 3).sum())
        within_7_days = int((self.df["Diff_BackSchedule"].abs() <= 7).sum())
        
        mean_diff = float(self.df["Diff_BackSchedule"].mean())
        std_diff = float(self.df["Diff_BackSchedule"].std())
        median_diff = float(self.df["Diff_BackSchedule"].median())
        
        early_deliveries = int((self.df["Diff_Ship_Req"] < 0).sum())
        late_deliveries = int((self.df["Diff_Ship_Req"] > 0).sum())
        on_time_deliveries = int((self.df["Diff_Ship_Req"] == 0).sum())
        
        quality_score = (
            (exact_match * 1.0 + 
             (within_1_day - exact_match) * 0.8 + 
             (within_3_days - within_1_day) * 0.6 + 
             (within_7_days - within_3_days) * 0.3) / total * 100
        )
        
        self.metrics = {
            "total_records": total,
            "exact_match_pct": round(100 * exact_match / total, 2),
            "within_1_day_pct": round(100 * within_1_day / total, 2),
            "within_3_days_pct": round(100 * within_3_days / total, 2),
            "within_7_days_pct": round(100 * within_7_days / total, 2),
            "mean_diff": round(mean_diff, 2),
            "std_diff": round(std_diff, 2),
            "median_diff": round(median_diff, 2),
            "quality_score": round(quality_score, 1),
            "early_deliveries_pct": round(100 * early_deliveries / total, 2),
            "late_deliveries_pct": round(100 * late_deliveries / total, 2),
            "on_time_deliveries_pct": round(100 * on_time_deliveries / total, 2)
        }
    
    def create_gaussian_chart(self):
        """Crear gráfico de campana de Gauss CORREGIDO: ShipDate - ReqDate vs ShipDate - ReqDate_Original"""
        if self.df is None or len(self.df) == 0:
            return None
            
        try:
            import matplotlib
            matplotlib.use('Agg')
            
            plt.style.use('dark_background')
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 8))
            fig.patch.set_facecolor(COLORS['surface'])
            
            # CAMPANA 1: BackSchedule (ShipDate - ReqDate) - CORREGIDO
            backschedule_data = self.df["Diff_Ship_Req"].dropna()  # CAMBIADO: Ahora usa ShipDate - ReqDate
            
            if len(backschedule_data) > 0:
                # Filtrar outliers para mejor visualización
                q1, q3 = backschedule_data.quantile([0.25, 0.75])
                iqr = q3 - q1
                lower_bound = q1 - 3 * iqr
                upper_bound = q3 + 3 * iqr
                filtered_data = backschedule_data[
                    (backschedule_data >= lower_bound) & 
                    (backschedule_data <= upper_bound)
                ]
                
                # Histograma con más bins para detalle
                n, bins, patches = ax1.hist(filtered_data, bins=80, alpha=0.7, 
                                            density=True, edgecolor='white', linewidth=0.3)
                
                # Colorear las barras: azul como en la imagen original
                for i, patch in enumerate(patches):
                    patch.set_facecolor('#1E90FF')  # DodgerBlue como en la imagen
                
                # Curva normal roja como en la imagen original
                mean_val = np.mean(filtered_data)
                std_val = np.std(filtered_data)
                x = np.linspace(filtered_data.min(), filtered_data.max(), 300)
                
                gaussian = (1 / (std_val * math.sqrt(2 * math.pi))) * np.exp(-0.5 * ((x - mean_val) / std_val) ** 2)
                ax1.plot(x, gaussian, 'red', linewidth=3, 
                        label=f'Normal(μ={mean_val:.1f}, σ={std_val:.1f})')
                
                # Línea vertical roja en cero como en la imagen original
                ax1.axvline(x=0, color='red', linestyle='-', linewidth=3, alpha=0.9)
                
                # Agregar texto con estadísticas - CORREGIDO
                negative_count = len(backschedule_data[backschedule_data < 0])  # PROBLEMAS (material llega después de necesitarlo)
                positive_count = len(backschedule_data[backschedule_data > 0])  # BUENO (material llega antes - BackSchedule correcto)
                zero_count = len(backschedule_data[backschedule_data == 0])     # JUSTO A TIEMPO
                total_count = len(backschedule_data)
                negative_pct = 100 * negative_count / total_count
                positive_pct = 100 * positive_count / total_count
                zero_pct = 100 * zero_count / total_count
                
                # Texto en esquina superior izquierda
                textstr = f'❌ Problemas (llega después): {negative_count:,} ({negative_pct:.1f}%)\n⭕ Justo a tiempo: {zero_count:,} ({zero_pct:.1f}%)\n✅ BackSchedule correcto (llega antes): {positive_count:,} ({positive_pct:.1f}%)\nTotal: {total_count:,}'
                props = dict(boxstyle='round', facecolor='black', alpha=0.8)
                ax1.text(0.02, 0.98, textstr, transform=ax1.transAxes, fontsize=10,
                        verticalalignment='top', bbox=props, color='white')
            
            # Estilo como en la imagen original
            ax1.set_title('BackSchedule: ShipDate - ReqDate\n(Positivo=Material llega ANTES ✅, Negativo=Material llega DESPUÉS ❌)', 
                            color='white', fontsize=16, fontweight='bold', pad=20)
            ax1.set_xlabel('Diferencia (días)', color='white', fontsize=14)
            ax1.set_ylabel('Densidad', color='white', fontsize=14)
            ax1.legend(loc='upper right', fontsize=12)
            ax1.grid(True, alpha=0.3, color='white')
            ax1.set_facecolor('#2C3E50')  # Fondo azul oscuro como en la imagen
            
            # Cambiar colores de los ejes
            ax1.tick_params(colors='white')
            ax1.spines['bottom'].set_color('white')
            ax1.spines['top'].set_color('white')
            ax1.spines['right'].set_color('white')
            ax1.spines['left'].set_color('white')
            
            # CAMPANA 2: Sistema Original (ShipDate - ReqDate_Original) 
            original_data = self.df["Diff_Ship_Original"].dropna()
            
            if len(original_data) > 0:
                # Filtrar outliers
                q1_orig, q3_orig = original_data.quantile([0.25, 0.75])
                iqr_orig = q3_orig - q1_orig
                lower_bound_orig = q1_orig - 3 * iqr_orig
                upper_bound_orig = q3_orig + 3 * iqr_orig
                filtered_orig = original_data[
                    (original_data >= lower_bound_orig) & 
                    (original_data <= upper_bound_orig)
                ]
                
                # Histograma 
                n2, bins2, patches2 = ax2.hist(filtered_orig, bins=80, alpha=0.7, 
                                                density=True, edgecolor='white', linewidth=0.3)
                
                # Colorear las barras: verde para diferenciar del BackSchedule
                for i, patch in enumerate(patches2):
                    patch.set_facecolor('#32CD32')  # LimeGreen
                
                # Curva normal roja
                mean_val2 = np.mean(filtered_orig)
                std_val2 = np.std(filtered_orig)
                x2 = np.linspace(filtered_orig.min(), filtered_orig.max(), 300)
                
                gaussian2 = (1 / (std_val2 * math.sqrt(2 * math.pi))) * np.exp(-0.5 * ((x2 - mean_val2) / std_val2) ** 2)
                ax2.plot(x2, gaussian2, 'red', linewidth=3,
                        label=f'Normal(μ={mean_val2:.1f}, σ={std_val2:.1f})')
                
                # Línea vertical roja en cero
                ax2.axvline(x=0, color='red', linestyle='-', linewidth=3, alpha=0.9)
                
                # Estadísticas
                negative_count2 = len(original_data[original_data < 0])  # PROBLEMAS (material llega después)
                positive_count2 = len(original_data[original_data > 0])  # BUENO (material llega antes)
                zero_count2 = len(original_data[original_data == 0])     # JUSTO A TIEMPO
                total_count2 = len(original_data)
                negative_pct2 = 100 * negative_count2 / total_count2
                positive_pct2 = 100 * positive_count2 / total_count2
                zero_pct2 = 100 * zero_count2 / total_count2
                
                textstr2 = f'❌ Problemas (llega después): {negative_count2:,} ({negative_pct2:.1f}%)\n⭕ Justo a tiempo: {zero_count2:,} ({zero_pct2:.1f}%)\n✅ BackSchedule correcto (llega antes): {positive_count2:,} ({positive_pct2:.1f}%)\nTotal: {total_count2:,}'
                props2 = dict(boxstyle='round', facecolor='black', alpha=0.8)
                ax2.text(0.02, 0.98, textstr2, transform=ax2.transAxes, fontsize=10,
                        verticalalignment='top', bbox=props2, color='white')
            
            ax2.set_title('Sistema Original: ShipDate - ReqDate_Original\n(Positivo=Material llega ANTES ✅, Negativo=Material llega DESPUÉS ❌)', 
                            color='white', fontsize=16, fontweight='bold', pad=20)
            ax2.set_xlabel('Diferencia (días)', color='white', fontsize=14)
            ax2.set_ylabel('Densidad', color='white', fontsize=14)
            ax2.legend(loc='upper right', fontsize=12)
            ax2.grid(True, alpha=0.3, color='white')
            ax2.set_facecolor('#2C3E50')
            
            # Cambiar colores de los ejes
            ax2.tick_params(colors='white')
            ax2.spines['bottom'].set_color('white')
            ax2.spines['top'].set_color('white')
            ax2.spines['right'].set_color('white')
            ax2.spines['left'].set_color('white')
            
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
            print(f"Error creando gráfico: {e}")
            return None
    
    def get_effectiveness_summary(self):
        """Obtener resumen de efectividad CORREGIDO"""
        if self.df is None or len(self.df) == 0:
            return {}
        
        # CORREGIDO: Ahora usamos las variables correctas
        backschedule_problems = len(self.df[self.df["Diff_Ship_Req"] < 0])        # Material llega DESPUÉS (PROBLEMAS)
        original_problems = len(self.df[self.df["Diff_Ship_Original"] < 0])       # Material llega DESPUÉS (PROBLEMAS)
        
        backschedule_good = len(self.df[self.df["Diff_Ship_Req"] > 0])            # Material llega ANTES (BUENO)
        original_good = len(self.df[self.df["Diff_Ship_Original"] > 0])           # Material llega ANTES (BUENO)
        
        backschedule_ontime = len(self.df[self.df["Diff_Ship_Req"] == 0])         # Material llega JUSTO A TIEMPO
        original_ontime = len(self.df[self.df["Diff_Ship_Original"] == 0])        # Material llega JUSTO A TIEMPO
        
        total = len(self.df)
        
        # Lead time promedio SOLO para casos con BackSchedule (valores positivos)
        bs_good = self.df[self.df["Diff_Ship_Req"] > 0]
        orig_good = self.df[self.df["Diff_Ship_Original"] > 0]
        
        bs_leadtime_avg = bs_good["Diff_Ship_Req"].mean() if len(bs_good) > 0 else 0
        orig_leadtime_avg = orig_good["Diff_Ship_Original"].mean() if len(orig_good) > 0 else 0
        
        return {
            "total_records": total,
            "backschedule_problems": backschedule_problems,
            "original_problems": original_problems,
            "backschedule_good": backschedule_good,
            "original_good": original_good,
            "backschedule_ontime": backschedule_ontime,
            "original_ontime": original_ontime,
            "backschedule_problems_pct": round(100 * backschedule_problems / total, 2),
            "original_problems_pct": round(100 * original_problems / total, 2),
            "difference": backschedule_problems - original_problems,
            "bs_leadtime_avg": round(bs_leadtime_avg, 1),
            "orig_leadtime_avg": round(orig_leadtime_avg, 1),
            "better_system": "BackSchedule" if backschedule_problems < original_problems 
                            else "Original" if original_problems < backschedule_problems 
                            else "Empate"
        }
    
    def export_effectiveness_analysis_to_excel(self, export_path=None):
        """Exportar análisis CORREGIDO con categorías separadas"""
        if self.df is None or len(self.df) == 0:
            return {"success": False, "message": "No hay datos para exportar"}
        
        try:
            # Separar por categorías - CORREGIDO
            backschedule_problems = self.df[self.df["Diff_Ship_Req"] < 0].copy()      # Material llega DESPUÉS (PROBLEMAS)
            backschedule_ontime = self.df[self.df["Diff_Ship_Req"] == 0].copy()       # Material llega JUSTO A TIEMPO
            backschedule_good = self.df[self.df["Diff_Ship_Req"] > 0].copy()          # Material llega ANTES (BUENO)
            
            original_problems = self.df[self.df["Diff_Ship_Original"] < 0].copy()     # Material llega DESPUÉS (PROBLEMAS)
            original_ontime = self.df[self.df["Diff_Ship_Original"] == 0].copy()      # Material llega JUSTO A TIEMPO
            original_good = self.df[self.df["Diff_Ship_Original"] > 0].copy()         # Material llega ANTES (BUENO)
            
            if export_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                export_path = f"Analisis_BackSchedule_Completo_{timestamp}.xlsx"
            
            # Conteos para debug
            bs_problems_count = len(backschedule_problems)
            bs_ontime_count = len(backschedule_ontime)
            bs_good_count = len(backschedule_good)
            
            orig_problems_count = len(original_problems)
            orig_ontime_count = len(original_ontime)
            orig_good_count = len(original_good)
            
            print(f"DEBUG - BS: Problemas={bs_problems_count}, JustoATiempo={bs_ontime_count}, Buenos={bs_good_count}")
            print(f"DEBUG - Orig: Problemas={orig_problems_count}, JustoATiempo={orig_ontime_count}, Buenos={orig_good_count}")
            
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                # SIEMPRE exportar todas las hojas, aunque estén vacías
                
                # BackSchedule - Problemas
                if len(backschedule_problems) > 0:
                    backschedule_problems.to_excel(writer, sheet_name='BS_Problemas', index=False)
                else:
                    # Crear hoja vacía con headers
                    empty_df = pd.DataFrame(columns=self.df.columns)
                    empty_df.to_excel(writer, sheet_name='BS_Problemas', index=False)
                
                # BackSchedule - Justo a Tiempo
                if len(backschedule_ontime) > 0:
                    backschedule_ontime.to_excel(writer, sheet_name='BS_JustoATiempo', index=False)
                else:
                    empty_df = pd.DataFrame(columns=self.df.columns)
                    empty_df.to_excel(writer, sheet_name='BS_JustoATiempo', index=False)
                
                
                # Original - Problemas
                if len(original_problems) > 0:
                    original_problems.to_excel(writer, sheet_name='Orig_Problemas', index=False)
                else:
                    empty_df = pd.DataFrame(columns=self.df.columns)
                    empty_df.to_excel(writer, sheet_name='Orig_Problemas', index=False)
                
                # Original - Justo a Tiempo
                if len(original_ontime) > 0:
                    original_ontime.to_excel(writer, sheet_name='Orig_JustoATiempo', index=False)
                else:
                    empty_df = pd.DataFrame(columns=self.df.columns)
                    empty_df.to_excel(writer, sheet_name='Orig_JustoATiempo', index=False)
                
                
                # Resumen comparativo - CORREGIDO: Solo datos estadísticos
            
            # Convertir a ruta absoluta
            import os
            full_path = os.path.abspath(export_path)
            
            return {
                "success": True,
                "message": f"Exportado: {os.path.basename(export_path)}",
                "path": full_path,  # Ruta completa para abrir
                "backschedule_problems": bs_problems_count,
                "backschedule_ontime": bs_ontime_count,
                "backschedule_good": bs_good_count,
                "original_problems": orig_problems_count,
                "original_ontime": orig_ontime_count,
                "original_good": orig_good_count,
                "better_system": "BackSchedule" if bs_problems_count < orig_problems_count else "Original"
            }
            
        except Exception as e:
            print(f"ERROR en export: {e}")
            return {"success": False, "message": f"Error: {str(e)}"}

    def load_fcst_alignment_data(self, db_path=r"C:\Users\J.Vazquez\Desktop\R4Database.db"):
        """Cargar y validar alineación de FCST con WO"""
        try:
            print(f"Conectando a R4Database: {db_path}")
            conn = sqlite3.connect(db_path)
            
            # Query para FCST con OpenQty > 0 y WO no vacío
            fcst_query = """
            SELECT Entity, Proj, AC, FcstNo, Description, 
                ItemNo, ReqDate, QtyFcst, OpenQty, WO
            FROM fcst 
            WHERE OpenQty > 0 AND WO IS NOT NULL AND WO != ""
            """
            
            # Query para WOInquiry
            wo_query = """
            SELECT WONo, SO_FCST, Sub, ParentWO,DueDt
            FROM WOInquiry
            """
            
            # Cargar datos
            fcst_df = pd.read_sql_query(fcst_query, conn)
            wo_df = pd.read_sql_query(wo_query, conn)
            conn.close()
            
            print(f"✅ FCST cargado: {len(fcst_df)} registros")
            print(f"✅ WOInquiry cargado: {len(wo_df)} registros")
            
            # Procesar fechas
            fcst_df["ReqDate"] = pd.to_datetime(fcst_df["ReqDate"], errors='coerce')
            wo_df["DueDt"] = pd.to_datetime(wo_df["DueDt"], errors='coerce')
            
            # Cruzar tablas: FCST.WO = WOInquiry.WONo
            merged_df = fcst_df.merge(wo_df, left_on='WO', right_on='WONo', how='left', suffixes=('_fcst', '_wo'))
            
            print(f"✅ Registros cruzados: {len(merged_df)}")
            
            # Validar alineación: ReqDate vs DueDt
            merged_df = merged_df.dropna(subset=['ReqDate', 'DueDt'])
            merged_df['Date_Diff_Days'] = (merged_df['DueDt'] - merged_df['ReqDate']).dt.days
            merged_df['Is_Aligned'] = merged_df['Date_Diff_Days'] == 0
            
            # Métricas de alineación
            total_records = len(merged_df)
            aligned_count = merged_df['Is_Aligned'].sum()
            misaligned_count = total_records - aligned_count
            alignment_pct = (aligned_count / total_records * 100) if total_records > 0 else 0
            
            # Análisis de desalineación
            misaligned_df = merged_df[~merged_df['Is_Aligned']].copy()
            
            # Estadísticas de diferencias
            mean_diff = merged_df['Date_Diff_Days'].mean()
            std_diff = merged_df['Date_Diff_Days'].std()
            median_diff = merged_df['Date_Diff_Days'].median()
            
            # Guardar resultados
            self.fcst_data = merged_df
            self.fcst_metrics = {
                'total_records': total_records,
                'aligned_count': int(aligned_count),
                'misaligned_count': int(misaligned_count),
                'alignment_pct': round(alignment_pct, 2),
                'mean_diff_days': round(mean_diff, 2),
                'std_diff_days': round(std_diff, 2),
                'median_diff_days': round(median_diff, 2),
                'early_wo': int((merged_df['Date_Diff_Days'] < 0).sum()),  # DueDt antes que ReqDate
                'late_wo': int((merged_df['Date_Diff_Days'] > 0).sum()),   # DueDt después que ReqDate
            }
            
            return True
            
        except Exception as e:
            print(f"❌ Error cargando FCST: {e}")
            # Crear datos de ejemplo para testing
            self._create_sample_fcst_data()
            return False

    def load_sales_order_alignment_data(self, db_path=r"C:\Users\J.Vazquez\Desktop\R4Database.db"):
        """Cargar y validar alineación de Sales Orders con Work Orders"""
        try:
            print(f"Conectando a R4Database para SO-WO: {db_path}")
            conn = sqlite3.connect(db_path)
            
            # Query para Sales Orders con Opn_Q > 0 y WO_No no vacío
            so_query = """
            SELECT Entity, Proj, SO_No, Ln, Ord_Cd, Order_Dt, CustReqDt, Req_Dt, Pln_Ship, Prd_Dt,
                Cust, Cust_Name, AC, Item_Number, Description, Orig_Q, Opn_Q, WO_No, 
                Unit_Price, OpenValue, ST, Planner_Notes
            FROM sales_order_table
            WHERE Opn_Q > 0 AND WO_No IS NOT NULL AND WO_No != ""
            """
            
            # Query para WOInquiry  
            wo_query = """
            SELECT WONo, SO_FCST, Sub, ItemNumber, Description as WO_Description, 
                AC as WO_AC, DueDt, CompDt, CreateDt, St as WO_Status, OpnQ as WO_OpnQ, Srt
            FROM WOInquiry
            WHERE Srt IN ('GF', 'GFA', 'KZ')
            """
            
            # Cargar datos
            so_df = pd.read_sql_query(so_query, conn)
            wo_df = pd.read_sql_query(wo_query, conn)
            conn.close()
            
            print(f"✅ Sales Orders cargadas: {len(so_df)} registros")
            print(f"✅ WOInquiry cargado: {len(wo_df)} registros")
            
            # Crear llaves para el cruce
            # SO Key: SO_No + Ln
            so_df['SO_Key'] = so_df['SO_No'].astype(str) + '-' + so_df['Ln'].astype(str)
            
            # WO Key: SO_FCST + Sub
            wo_df['WO_Key'] = wo_df['SO_FCST'].astype(str) + '-' + wo_df['Sub'].astype(str)
            
            print(f"🔑 SO Keys únicas: {so_df['SO_Key'].nunique()}")
            print(f"🔑 WO Keys únicas: {wo_df['WO_Key'].nunique()}")
            
            # Cruzar tablas: SO_Key = WO_Key
            merged_df = so_df.merge(wo_df, left_on='SO_Key', right_on='WO_Key', how='left', suffixes=('_so', '_wo'))
            
            print(f"✅ Registros cruzados SO-WO: {len(merged_df)}")
            
            # Procesar fechas
            merged_df["Prd_Dt"] = pd.to_datetime(merged_df["Prd_Dt"], errors='coerce')
            merged_df["DueDt"] = pd.to_datetime(merged_df["DueDt"], errors='coerce')
            merged_df["Req_Dt"] = pd.to_datetime(merged_df["Req_Dt"], errors='coerce')
            merged_df["CustReqDt"] = pd.to_datetime(merged_df["CustReqDt"], errors='coerce')
            
            # Eliminar registros sin fechas válidas
            initial_count = len(merged_df)
            merged_df = merged_df.dropna(subset=['Prd_Dt', 'DueDt'])
            final_count = len(merged_df)
            
            print(f"📊 Registros con fechas válidas: {final_count}/{initial_count}")
            
            # Validar alineación: Prd_Dt (SO) vs DueDt (WO)
            merged_df['Date_Diff_Days'] = (merged_df['DueDt'] - merged_df['Prd_Dt']).dt.days
            merged_df['Is_Aligned'] = merged_df['Date_Diff_Days'] == 0
            
            # Categorizar diferencias
            merged_df['Alignment_Category'] = merged_df['Date_Diff_Days'].apply(
                lambda x: 'Perfecta' if x == 0 
                        else 'WO Temprana' if x < 0 
                        else 'WO Tardía'
            )
            
            # Métricas de alineación
            total_records = len(merged_df)
            aligned_count = int(merged_df['Is_Aligned'].sum())
            misaligned_count = total_records - aligned_count
            alignment_pct = (aligned_count / total_records * 100) if total_records > 0 else 0
            
            # Estadísticas de diferencias
            mean_diff = merged_df['Date_Diff_Days'].mean()
            std_diff = merged_df['Date_Diff_Days'].std()
            median_diff = merged_df['Date_Diff_Days'].median()
            
            # Conteos por categoría
            early_wo = int((merged_df['Date_Diff_Days'] < 0).sum())  # DueDt antes que Prd_Dt
            late_wo = int((merged_df['Date_Diff_Days'] > 0).sum())   # DueDt después que Prd_Dt
            
            # Guardar resultados
            self.so_data = merged_df
            self.so_metrics = {
                'total_records': total_records,
                'aligned_count': aligned_count,
                'misaligned_count': misaligned_count,
                'alignment_pct': round(alignment_pct, 2),
                'mean_diff_days': round(mean_diff, 2),
                'std_diff_days': round(std_diff, 2),
                'median_diff_days': round(median_diff, 2),
                'early_wo': early_wo,  # DueDt antes que Prd_Dt
                'late_wo': late_wo,    # DueDt después que Prd_Dt
                'so_keys_unique': so_df['SO_Key'].nunique(),
                'wo_keys_unique': wo_df['WO_Key'].nunique(),
                'matched_keys': int(merged_df['WO_Key'].notna().sum())
            }
            
            print(f"✅ SO-WO Analysis completado: {alignment_pct:.1f}% alineación")
            return True
            
        except Exception as e:
            print(f"❌ Error cargando SO-WO: {e}")
            # Crear datos de ejemplo para testing
            self._create_sample_so_data()
            return False

    def load_wo_materials_analysis(self, db_path=r"C:\Users\J.Vazquez\Desktop\R4Database.db"):
        """Cargar y analizar WO con/sin materiales (pr561)"""
        try:
            print(f"Conectando a R4Database para WO-Materials: {db_path}")
            print(f"¿Existe el archivo? {os.path.exists(db_path)}")
            conn = sqlite3.connect(db_path)
            print("✅ Conexión exitosa a la base de datos")
                    
            # Query para WOInquiry con filtro Srt y OpnQ > 0
            wo_query = """
            SELECT WONo, SO_FCST, Sub, ItemNumber, Description as WO_Description, 
                AC, DueDt, CompDt, CreateDt, St as WO_Status, OpnQ, Srt, Plnr, PlanType
            FROM WOInquiry
            WHERE Srt IN ('GF', 'GFA', 'KZ') AND OpnQ > 0
            """
            
            # Query para pr561 - todos los materiales
            pr561_query = """
            SELECT id, Entity, Project, ItemNo, FuseNo, Description, PlnType, Srt as PR_Srt, 
                St as PR_St, QtyOh, QtyIssue, QtyPending, ReqQty, ValQtyIss, ValNotIss, 
                ValRequired, WONo as PR_WONo, WODescripton, ReqDate
            FROM pr561
            """
            
            # Cargar datos
            print("🔄 Ejecutando query WOInquiry...")
            wo_df = pd.read_sql_query(wo_query, conn)
            print(f"✅ WOInquiry cargado: {len(wo_df)} registros")
            
            print("🔄 Ejecutando query PR561...")
            pr561_df = pd.read_sql_query(pr561_query, conn)
            print(f"✅ PR561 cargado: {len(pr561_df)} registros")
            conn.close()
            
            print("🔄 Agrupando materiales por WO...")
            # Agrupar materiales por WO para contar
            materials_summary = pr561_df.groupby('PR_WONo').agg({
                'id': 'count',  # Cantidad de materiales
                'ReqQty': 'sum',  # Total requerido
                'QtyIssue': 'sum',  # Total emitido
                'QtyPending': 'sum',  # Total pendiente
                'ValRequired': 'sum'  # Valor total requerido
            }).rename(columns={
                'id': 'Materials_Count',
                'ReqQty': 'Total_ReqQty',
                'QtyIssue': 'Total_QtyIssue', 
                'QtyPending': 'Total_QtyPending',
                'ValRequired': 'Total_ValRequired'
            }).reset_index()
            
            print(f"✅ WO únicas en PR561: {len(materials_summary)}")
            
            print("🔄 Cruzando WO con materiales...")
            # Cruzar WO con materiales (LEFT JOIN para incluir WO sin materiales)
            merged_df = wo_df.merge(materials_summary, left_on='WONo', right_on='PR_WONo', how='left')
            print(f"✅ Merge completado: {len(merged_df)} registros")
            
            print("🔄 Procesando campos...")
            # Identificar WO sin materiales (Materials_Count es NaN)
            merged_df['Has_Materials'] = merged_df['Materials_Count'].notna()
            merged_df['Materials_Count'] = merged_df['Materials_Count'].fillna(0).astype(int)
            merged_df['Total_ReqQty'] = merged_df['Total_ReqQty'].fillna(0)
            merged_df['Total_QtyIssue'] = merged_df['Total_QtyIssue'].fillna(0)
            merged_df['Total_QtyPending'] = merged_df['Total_QtyPending'].fillna(0)
            merged_df['Total_ValRequired'] = merged_df['Total_ValRequired'].fillna(0)
            
            # Categorizar WO
            merged_df['WO_Category'] = merged_df.apply(
                lambda row: 'Con Materiales' if row['Has_Materials'] 
                        else 'Sin Materiales', axis=1
            )
            
            print("🔄 Procesando fechas...")
            # Procesar fechas
            merged_df['DueDt'] = pd.to_datetime(merged_df['DueDt'], errors='coerce')
            merged_df['CompDt'] = pd.to_datetime(merged_df['CompDt'], errors='coerce')
            merged_df['CreateDt'] = pd.to_datetime(merged_df['CreateDt'], errors='coerce')
            
            print("🔄 Calculando métricas generales...")
            # Métricas generales
            total_wo = len(merged_df)
            wo_with_materials = int(merged_df['Has_Materials'].sum())
            wo_without_materials = total_wo - wo_with_materials
            materials_pct = (wo_with_materials / total_wo * 100) if total_wo > 0 else 0
            
            print(f"📊 Métricas básicas calculadas:")
            print(f"   Total WO: {total_wo}")
            print(f"   Con materiales: {wo_with_materials}")
            print(f"   Sin materiales: {wo_without_materials}")
            print(f"   % con materiales: {materials_pct}")
            
            print("🔄 Calculando métricas por Srt...")
            # Métricas por tipo Srt - MÉTODO CORREGIDO
            srt_totals = merged_df['Srt'].value_counts()
            
            # Calcular sin materiales por Srt de forma directa y robusta
            srt_gf_total = len(merged_df[merged_df['Srt'] == 'GF'])
            srt_gfa_total = len(merged_df[merged_df['Srt'] == 'GFA'])
            srt_kz_total = len(merged_df[merged_df['Srt'] == 'KZ'])
            
            srt_gf_without = len(merged_df[(merged_df['Srt'] == 'GF') & (~merged_df['Has_Materials'])])
            srt_gfa_without = len(merged_df[(merged_df['Srt'] == 'GFA') & (~merged_df['Has_Materials'])])
            srt_kz_without = len(merged_df[(merged_df['Srt'] == 'KZ') & (~merged_df['Has_Materials'])])
            
            print(f"📊 Análisis Srt CORREGIDO:")
            print(f"   GF: {srt_gf_total} total, {srt_gf_without} sin materiales")
            print(f"   GFA: {srt_gfa_total} total, {srt_gfa_without} sin materiales") 
            print(f"   KZ: {srt_kz_total} total, {srt_kz_without} sin materiales")
            print(f"   SUMA sin materiales: {srt_gf_without + srt_gfa_without + srt_kz_without}")
            
            print("🔄 Calculando métricas por WO_Status...")
            # Métricas por WO_Status
            status_totals = merged_df['WO_Status'].value_counts()
            
            print(f"📊 Análisis WO_Status:")
            print(f"   status_totals: {dict(status_totals)}")
            
            # Crear métricas dinámicas por cada status
            status_metrics = {}
            for status in status_totals.index:
                total_for_status = len(merged_df[merged_df['WO_Status'] == status])
                without_for_status = len(merged_df[(merged_df['WO_Status'] == status) & (~merged_df['Has_Materials'])])
                with_for_status = total_for_status - without_for_status
                
                status_metrics[f'status_{status}_total'] = total_for_status
                status_metrics[f'status_{status}_with'] = with_for_status
                status_metrics[f'status_{status}_without'] = without_for_status
                status_metrics[f'status_{status}_without_pct'] = round(100 * without_for_status / max(total_for_status, 1), 1)
                
                print(f"   Status {status}: {total_for_status} total, {without_for_status} sin materiales ({100 * without_for_status / max(total_for_status, 1):.1f}%)")

            print(f"📊 Métricas por Status creadas: {len(status_metrics)} métricas")
            
            print("🔄 Calculando valores promedio...")
            # Análisis de valor para WO con materiales - CORREGIDO
            wo_with_mat = merged_df[merged_df['Has_Materials']]
            
            print(f"   WO con materiales: {len(wo_with_mat)}")
            
            if len(wo_with_mat) > 0:
                try:
                    print("   🔍 Calculando avg_materials_per_wo...")
                    avg_materials_per_wo = float(wo_with_mat['Materials_Count'].mean())
                    print(f"   avg_materials_per_wo: {avg_materials_per_wo}")
                    
                    print("   🔍 Calculando total_value_required...")
                    total_value_required = float(wo_with_mat['Total_ValRequired'].sum())
                    print(f"   total_value_required: {total_value_required}")
                    
                except Exception as calc_error:
                    print(f"   ❌ Error en cálculo: {calc_error}")
                    avg_materials_per_wo = 0.0
                    total_value_required = 0.0
            else:
                print("   ⚠️ No hay WO con materiales")
                avg_materials_per_wo = 0.0
                total_value_required = 0.0
            
            print("🔄 Construyendo diccionario de métricas...")
            # Guardar resultados
            self.wo_materials_data = merged_df
            self.wo_materials_raw = pr561_df
            
            # Construir métricas paso a paso para identificar el error
            metrics = {}
            
            try:
                print("   ✅ Agregando métricas básicas...")
                metrics['total_wo'] = total_wo
                metrics['wo_with_materials'] = wo_with_materials
                metrics['wo_without_materials'] = wo_without_materials
                
                print("   ✅ Agregando porcentajes...")
                metrics['materials_pct'] = round(materials_pct, 2)
                metrics['without_materials_pct'] = round(100 - materials_pct, 2)
                
                print("   ✅ Agregando promedios...")
                metrics['avg_materials_per_wo'] = round(avg_materials_per_wo, 1)
                metrics['total_value_required'] = round(total_value_required, 2)
                
                print("   ✅ Agregando métricas Srt CORREGIDAS...")
                metrics['srt_gf_total'] = srt_gf_total
                metrics['srt_gfa_total'] = srt_gfa_total
                metrics['srt_kz_total'] = srt_kz_total
                metrics['srt_gf_without'] = srt_gf_without
                metrics['srt_gfa_without'] = srt_gfa_without
                metrics['srt_kz_without'] = srt_kz_without
                
                print("   ✅ Agregando métricas de Status...")
                # Agregar métricas de status
                metrics.update(status_metrics)
                
            except Exception as metrics_error:
                print(f"   ❌ Error construyendo métricas: {metrics_error}")
                print(f"   ❌ Error en línea: {metrics_error.__traceback__.tb_lineno}")
                raise metrics_error
            
            self.wo_materials_metrics = metrics
            
            # VERIFICACIÓN FINAL
            print(f"\n🔍 VERIFICACIÓN FINAL:")
            print(f"   📊 Total WO: {total_wo}")
            print(f"   ✅ Con materiales: {wo_with_materials} ({materials_pct:.1f}%)")
            print(f"   ❌ Sin materiales: {wo_without_materials} ({100-materials_pct:.1f}%)")
            print(f"   🔧 Por Srt sin materiales: GF={srt_gf_without}, GFA={srt_gfa_without}, KZ={srt_kz_without}")
            print(f"   🧮 Suma Srt: {srt_gf_without + srt_gfa_without + srt_kz_without} (debe coincidir con {wo_without_materials})")
            
            return True
            
        except Exception as e:
            print(f"❌ Error cargando WO-Materials: {e}")
            print(f"❌ Error tipo: {type(e)}")
            print(f"❌ Error args: {e.args}")
            import traceback
            print(f"❌ Traceback completo:")
            traceback.print_exc()
            
            # Crear datos de ejemplo para testing
            self._create_sample_wo_materials_data()
            return False

    def _create_sample_fcst_data(self):
        """Crear datos de ejemplo para FCST"""
        np.random.seed(42)
        n_records = 150
        
        # Datos simulados
        base_date = datetime(2025, 8, 1)
        wo_numbers = [f"WO-{1000+i:04d}" for i in range(50)]
        
        data = []
        for i in range(n_records):
            req_date = base_date + timedelta(days=i + np.random.randint(-10, 30))
            
            # Simular desalineación: 70% alineado, 30% desalineado
            if np.random.random() < 0.7:
                due_dt = req_date  # Perfectamente alineado
            else:
                # Desalineado con variaciones de -10 a +15 días
                diff_days = np.random.choice([-10, -5, -3, -1, 1, 3, 5, 10, 15], 
                                        p=[0.05, 0.10, 0.15, 0.20, 0.20, 0.15, 0.10, 0.03, 0.02])
                due_dt = req_date + timedelta(days=diff_days)
            
            record = {
                'Entity_fcst': f'Entity_{i%3}',
                'WO': np.random.choice(wo_numbers),
                'ItemNo': f'Item-{100+i:03d}',
                'Description_fcst': f'FCST Item {i}',
                'ReqDate': req_date,
                'OpenQty': np.random.randint(1, 100),
                'WONo': np.random.choice(wo_numbers),
                'DueDt': due_dt,
                'Description_wo': f'WO Item {i}',
                'AC_wo': f'AC{i%3}',
                'Date_Diff_Days': (due_dt - req_date).days,
                'Is_Aligned': (due_dt - req_date).days == 0
            }
            data.append(record)
        
        self.fcst_data = pd.DataFrame(data)
        
        # Calcular métricas
        total = len(self.fcst_data)
        aligned = self.fcst_data['Is_Aligned'].sum()
        
        self.fcst_metrics = {
            'total_records': total,
            'aligned_count': int(aligned),
            'misaligned_count': int(total - aligned),
            'alignment_pct': round(100 * aligned / total, 2),
            'mean_diff_days': round(self.fcst_data['Date_Diff_Days'].mean(), 2),
            'std_diff_days': round(self.fcst_data['Date_Diff_Days'].std(), 2),
            'median_diff_days': round(self.fcst_data['Date_Diff_Days'].median(), 2),
            'early_wo': int((self.fcst_data['Date_Diff_Days'] < 0).sum()),
            'late_wo': int((self.fcst_data['Date_Diff_Days'] > 0).sum()),
        }

    def _create_sample_so_data(self):
        """Crear datos de ejemplo para SO-WO"""
        np.random.seed(42)
        n_records = 200
        
        # Datos simulados
        base_date = datetime(2025, 8, 1)
        so_numbers = [f"SO-{2000+i:04d}" for i in range(50)]
        
        data = []
        for i in range(n_records):
            so_no = np.random.choice(so_numbers)
            ln = np.random.randint(1, 5)
            so_key = f"{so_no}-{ln}"
            
            prd_dt = base_date + timedelta(days=i + np.random.randint(-10, 30))
            
            # Simular alineación: 65% alineado, 35% desalineado
            if np.random.random() < 0.65:
                due_dt = prd_dt  # Perfectamente alineado
            else:
                # Desalineado con variaciones de -15 a +10 días
                diff_days = np.random.choice([-15, -7, -3, -1, 1, 3, 7, 10], 
                                        p=[0.05, 0.15, 0.20, 0.25, 0.20, 0.10, 0.03, 0.02])
                due_dt = prd_dt + timedelta(days=diff_days)
            
            record = {
                'Entity': f'Entity_{i%3}',
                'SO_No': so_no,
                'Ln': ln,
                'SO_Key': so_key,
                'Item_Number': f'Item-SO-{100+i:03d}',
                'Description_so': f'SO Item {i}',
                'Prd_Dt': prd_dt,
                'Opn_Q': np.random.randint(1, 100),
                'WO_No': f"WO-SO-{1000+i:04d}",
                'AC': f'AC{i%3}',
                'Cust': f'Customer_{i%10}',
                'OpenValue': np.random.randint(1000, 50000),
                'WO_Key': so_key,  # Para simular match
                'DueDt': due_dt,
                'WO_Description': f'WO for SO Item {i}',
                'WO_AC': f'AC{i%3}',
                'WO_OpnQ': np.random.randint(1, 100),
                'Date_Diff_Days': (due_dt - prd_dt).days,
                'Is_Aligned': (due_dt - prd_dt).days == 0
            }
            data.append(record)
        
        self.so_data = pd.DataFrame(data)
        
        # Calcular métricas
        total = len(self.so_data)
        aligned = self.so_data['Is_Aligned'].sum()
        
        self.so_metrics = {
            'total_records': total,
            'aligned_count': int(aligned),
            'misaligned_count': int(total - aligned),
            'alignment_pct': round(100 * aligned / total, 2),
            'mean_diff_days': round(self.so_data['Date_Diff_Days'].mean(), 2),
            'std_diff_days': round(self.so_data['Date_Diff_Days'].std(), 2),
            'median_diff_days': round(self.so_data['Date_Diff_Days'].median(), 2),
            'early_wo': int((self.so_data['Date_Diff_Days'] < 0).sum()),
            'late_wo': int((self.so_data['Date_Diff_Days'] > 0).sum()),
            'so_keys_unique': len(so_numbers) * 4,  # Simulado
            'wo_keys_unique': len(so_numbers) * 4,  # Simulado
            'matched_keys': total
        }

    def _create_sample_wo_materials_data(self):
        """Crear datos de ejemplo para WO-Materials"""
        np.random.seed(42)
        n_wo = 150
        
        # Datos simulados para WO
        base_date = datetime(2025, 8, 1)
        wo_numbers = [f"WO-MAT-{3000+i:04d}" for i in range(n_wo)]
        srt_types = ['GF', 'GFA', 'KZ']
        
        wo_data = []
        pr561_data = []
        
        for i, wo_no in enumerate(wo_numbers):
            # Datos de WO
            srt = np.random.choice(srt_types, p=[0.5, 0.3, 0.2])  # GF más común
            due_dt = base_date + timedelta(days=i + np.random.randint(-5, 20))
            
            wo_record = {
                'WONo': wo_no,
                'SO_FCST': f'SO-{2000+i//3:04d}',
                'Sub': str(i % 5 + 1),
                'ItemNumber': f'Item-WO-{100+i:03d}',
                'WO_Description': f'WO Material Item {i}',
                'AC': f'AC{i%3}',
                'DueDt': due_dt,
                'CompDt': None,
                'CreateDt': due_dt - timedelta(days=np.random.randint(5, 15)),
                'WO_Status': np.random.choice(['O', 'R', 'C'], p=[0.7, 0.2, 0.1]),
                'OpnQ': np.random.randint(1, 100),
                'Srt': srt,
                'Plnr': f'PLN{i%5}',
                'PlanType': 'MTO'
            }
            wo_data.append(wo_record)
            
            # Simular materiales: 75% de WO tienen materiales, 25% no tienen
            has_materials = np.random.random() < 0.75
            
            if has_materials:
                # Generar 1-5 materiales por WO
                n_materials = np.random.randint(1, 6)
                for j in range(n_materials):
                    material_record = {
                        'id': i * 10 + j,
                        'Entity': f'Entity_{i%3}',
                        'Project': f'Proj_{i%5}',
                        'ItemNo': f'MAT-{200+j:03d}',
                        'FuseNo': f'FUSE-{i}-{j}',
                        'Description': f'Material {j} for WO {wo_no}',
                        'PlnType': 'BUY',
                        'PR_Srt': np.random.choice(['P', 'O', 'I']),
                        'PR_St': np.random.choice(['O', 'C', 'R']),
                        'QtyOh': np.random.randint(0, 50),
                        'QtyIssue': np.random.randint(0, 20),
                        'QtyPending': np.random.randint(0, 30),
                        'ReqQty': np.random.randint(10, 100),
                        'ValQtyIss': np.random.uniform(100, 1000),
                        'ValNotIss': np.random.uniform(50, 500),
                        'ValRequired': np.random.uniform(200, 1500),
                        'PR_WONo': wo_no,
                        'WODescripton': f'WO Material Item {i}',
                        'ReqDate': due_dt - timedelta(days=np.random.randint(1, 10))
                    }
                    pr561_data.append(material_record)
        
        # Crear DataFrames
        wo_df = pd.DataFrame(wo_data)
        pr561_df = pd.DataFrame(pr561_data)
        
        # Agrupar materiales por WO
        materials_summary = pr561_df.groupby('PR_WONo').agg({
            'id': 'count',
            'ReqQty': 'sum',
            'QtyIssue': 'sum',
            'QtyPending': 'sum',
            'ValRequired': 'sum'
        }).rename(columns={
            'id': 'Materials_Count',
            'ReqQty': 'Total_ReqQty',
            'QtyIssue': 'Total_QtyIssue',
            'QtyPending': 'Total_QtyPending',
            'ValRequired': 'Total_ValRequired'
        }).reset_index()
        
        # Cruzar datos
        merged_df = wo_df.merge(materials_summary, left_on='WONo', right_on='PR_WONo', how='left')
        
        # Procesar campos
        merged_df['Has_Materials'] = merged_df['Materials_Count'].notna()
        merged_df['Materials_Count'] = merged_df['Materials_Count'].fillna(0).astype(int)
        merged_df['Total_ReqQty'] = merged_df['Total_ReqQty'].fillna(0)
        merged_df['Total_QtyIssue'] = merged_df['Total_QtyIssue'].fillna(0)
        merged_df['Total_QtyPending'] = merged_df['Total_QtyPending'].fillna(0)
        merged_df['Total_ValRequired'] = merged_df['Total_ValRequired'].fillna(0)
        
        merged_df['WO_Category'] = merged_df.apply(
            lambda row: 'Con Materiales' if row['Has_Materials'] else 'Sin Materiales', axis=1
        )
        
        # Calcular métricas
        total_wo = len(merged_df)
        wo_with_materials = int(merged_df['Has_Materials'].sum())
        wo_without_materials = total_wo - wo_with_materials
        materials_pct = (wo_with_materials / total_wo * 100)
        
        # Métricas por Srt
        srt_counts = merged_df['Srt'].value_counts()
        wo_without_by_srt = merged_df[~merged_df['Has_Materials']]['Srt'].value_counts()
        
        # Guardar datos
        self.wo_materials_data = merged_df
        self.wo_materials_raw = pr561_df
        self.wo_materials_metrics = {
            'total_wo': total_wo,
            'wo_with_materials': wo_with_materials,
            'wo_without_materials': wo_without_materials,
            'materials_pct': round(materials_pct, 2),
            'without_materials_pct': round(100 - materials_pct, 2),
            'avg_materials_per_wo': round(merged_df[merged_df['Has_Materials']]['Materials_Count'].mean(), 1),
            'total_value_required': round(merged_df['Total_ValRequired'].sum(), 2),
            'srt_gf_total': int(srt_counts.get('GF', 0)),
            'srt_gfa_total': int(srt_counts.get('GFA', 0)),
            'srt_kz_total': int(srt_counts.get('KZ', 0)),
            'srt_gf_without': int(wo_without_by_srt.get('GF', 0)),
            'srt_gfa_without': int(wo_without_by_srt.get('GFA', 0)),
            'srt_kz_without': int(wo_without_by_srt.get('KZ', 0)),
        }
        
        print(f"🧪 Datos de ejemplo WO-Materials creados:")
        print(f"   📊 Total WO: {total_wo}")
        print(f"   ✅ Con materiales: {wo_with_materials} ({materials_pct:.1f}%)")
        print(f"   ❌ Sin materiales: {wo_without_materials} ({100-materials_pct:.1f}%)")

    """aqui se agrego"""

    def load_obsolete_expired_neteable_analysis(self, db_path=r"C:\Users\J.Vazquez\Desktop\R4Database.db"):
        """Cargar y analizar materiales obsoletos/expirados en localidades neteables"""
        try:
            print(f"Conectando a R4Database para análisis OBS/EXP LOC NET: {db_path}")
            conn = sqlite3.connect(db_path)
            
            # Query para localidades neteables
            locations_query = """
            SELECT id, Whs, BinID, LocID, Description, ExclBin, Zone, Container, 
                TagNetable, Section, Type
            FROM whs_location_in36851
            WHERE UPPER(TRIM(TagNetable)) IN ('YES', 'Y', 'TRUE', '1', 'SI', 'SÍ')
            """
            
            # Query para materiales con inventario
            materials_query = """
            SELECT id, Wh, ItemNo, Description, UM, Prod, ExpDate, Rec, QtyOH, 
                Bin, Lot, OrderBy, LRev, IRev, PlanType, EntCode, Proj, 
                ManufDate, Shelf_Life
            FROM in521
            WHERE QtyOH > 0
            """
            
            # Cargar datos
            print("🔄 Ejecutando query whs_location_in36851...")
            locations_df = pd.read_sql_query(locations_query, conn)
            print(f"✅ Localidades neteables cargadas: {len(locations_df)} registros")
            
            print("🔄 Ejecutando query in521...")
            materials_df = pd.read_sql_query(materials_query, conn)
            print(f"✅ Materiales cargados: {len(materials_df)} registros")
            conn.close()
            
            # Limpiar y convertir a mayúsculas - MEJORADO
            locations_df['BinID'] = locations_df['BinID'].astype(str).str.upper().str.strip()
            locations_df['Whs'] = locations_df['Whs'].astype(str).str.upper().str.strip()
            locations_df['Zone'] = locations_df['Zone'].astype(str).str.upper().str.strip()

            materials_df['Bin'] = materials_df['Bin'].astype(str).str.upper().str.strip()
            materials_df['Wh'] = materials_df['Wh'].astype(str).str.upper().str.strip()

            # Validar TagNetable
            locations_df['TagNetable_Clean'] = locations_df['TagNetable'].astype(str).str.upper().str.strip()
            locations_df = locations_df[locations_df['TagNetable_Clean'].isin(['YES', 'Y', 'TRUE', '1', 'SI', 'SÍ'])]

            print(f"🔍 Validación TagNetable:")
            print(f"   Valores únicos encontrados: {locations_df['TagNetable'].unique()}")
            print(f"   Localidades válidas después del filtro: {len(locations_df)}")
            print(f"🔑 BinID únicas en localidades: {locations_df['BinID'].nunique()}")
            print(f"🔑 Bin únicas en materiales: {materials_df['Bin'].nunique()}")
            
            # Cruzar datos
            merged_df = materials_df.merge(
                locations_df[['BinID', 'TagNetable', 'Zone', 'Type', 'Description']], 
                left_on='Bin', 
                right_on='BinID', 
                how='inner',
                suffixes=('_material', '_location')
            )
            
            print(f"✅ Registros cruzados (materiales en localidades neteables): {len(merged_df)}")
            
            # Normalizar campos después del merge
            merged_df['Wh'] = merged_df['Wh'].astype(str).str.upper().str.strip()
            merged_df['Zone'] = merged_df['Zone'].astype(str).str.upper().str.strip()

            print(f"🔧 Después de normalizar:")
            print(f"   Warehouses únicos: {sorted(merged_df['Wh'].unique())}")
            print(f"   Zonas únicas: {sorted(merged_df['Zone'].unique())}")
            
            # CORREGIR: Limpiar Shelf_Life para remover "%" y convertir a número
            print(f"🔧 Limpiando datos de Shelf_Life...")
            
            # Limpiar Shelf_Life: remover "%" y convertir a número
            merged_df['Shelf_Life_Original'] = merged_df['Shelf_Life'].astype(str)
            merged_df['Shelf_Life_Clean'] = merged_df['Shelf_Life_Original'].str.replace('%', '').str.replace('nan', '')
            merged_df['Shelf_Life'] = pd.to_numeric(merged_df['Shelf_Life_Clean'], errors='coerce')
            
            # CORREGIR: Limpiar QtyOH para convertir a número
            merged_df['QtyOH'] = pd.to_numeric(merged_df['QtyOH'], errors='coerce')
            
            print(f"📊 Después de limpiar datos:")
            print(f"   Shelf_Life válidos: {merged_df['Shelf_Life'].notna().sum()}")
            print(f"   Shelf_Life rango: {merged_df['Shelf_Life'].min():.2f}% a {merged_df['Shelf_Life'].max():.2f}%")
            print(f"   QtyOH válidos: {merged_df['QtyOH'].notna().sum()}")
            
            # Procesar fechas
            merged_df['ExpDate'] = pd.to_datetime(merged_df['ExpDate'], errors='coerce')
            merged_df['ManufDate'] = pd.to_datetime(merged_df['ManufDate'], errors='coerce')
            
            initial_count = len(merged_df)

            # Distribución por rangos de porcentaje
            zero_shelf = len(merged_df[merged_df['Shelf_Life'] == 0])     # = 0% (PROBLEMA)
            expired_5 = len(merged_df[merged_df['Shelf_Life'] <= 5])     # <= 5%
            expired_10 = len(merged_df[merged_df['Shelf_Life'] <= 10])   # <= 10%
            expired_25 = len(merged_df[merged_df['Shelf_Life'] <= 25])   # <= 25%
            vigente_count = len(merged_df[merged_df['Shelf_Life'] > 25]) # > 25%
            null_count = merged_df['Shelf_Life'].isnull().sum()

            print(f"📈 Distribución Shelf_Life (%) - PARA NETEABLE:")
            print(f"   = 0% (PROBLEMA - debe moverse): {zero_shelf}")
            print(f"   <= 5% (Crítico): {expired_5 - zero_shelf}")
            print(f"   <= 10% (Muy Bajo): {expired_10 - expired_5}")
            print(f"   <= 25% (Bajo): {expired_25 - expired_10}")
            print(f"   > 25% (Aceptable en neteable): {vigente_count}")
            print(f"   Nulos/Inválidos: {null_count}")

            # CORREGIDO: SOLO materiales con Shelf_Life = 0% (estos deben moverse de localidades neteables)
            expired_df = merged_df[
                (merged_df['Shelf_Life'].notna()) & 
                (merged_df['Shelf_Life'] == 0)
            ].copy()
            expired_count = len(expired_df)

            print(f"🔍 DEBUG Filtro NETEABLE:")
            print(f"   Materiales con Shelf_Life válido: {merged_df['Shelf_Life'].notna().sum()}")
            print(f"   Materiales con Shelf_Life = 0% (deben moverse): {len(expired_df)}")
            if len(expired_df) > 0:
                print(f"   Todos los materiales filtrados tienen Shelf_Life = 0%")

            print(f"📊 RESULTADOS FINALES - KPI NETEABLE:")
            print(f"   Materiales totales en loc. neteables: {initial_count}")
            print(f"   🚨 Materiales con 0% en loc. neteables (DEBEN MOVERSE): {expired_count}")
            print(f"   Porcentaje que debe moverse: {100*expired_count/max(initial_count,1):.2f}%")

            # Categorizar - solo habrá una categoría ya que todos son 0%
            expired_df['Expiry_Category'] = 'Material Expirado (0%) - Debe moverse de localidad neteable'
            
            # Categorizar por cantidad
            expired_df['Qty_Category'] = expired_df['QtyOH'].apply(
                lambda x: 'Alto (>1000)' if pd.notna(x) and x > 1000
                        else 'Medio (100-1000)' if pd.notna(x) and x > 100
                        else 'Bajo (1-100)' if pd.notna(x) and x > 0
                        else 'Sin Stock'
            )
            
            # Categorizar por prioridad de movimiento
            expired_df['Move_Priority'] = expired_df['QtyOH'].apply(
                lambda x: 'Prioridad Alta (>1000 unidades)' if pd.notna(x) and x > 1000
                        else 'Prioridad Media (100-1000 unidades)' if pd.notna(x) and x > 100
                        else 'Prioridad Baja (<100 unidades)' if pd.notna(x) and x > 0
                        else 'Sin Stock'
            )
            
            # Calcular métricas
            total_materials_neteable = initial_count
            total_expired = expired_count
            expired_percentage = (expired_count / max(initial_count, 1) * 100)
            
            # Métricas por categoría
            expiry_counts = expired_df['Expiry_Category'].value_counts()
            zone_counts = expired_df['Zone'].value_counts()
            wh_counts = expired_df['Wh'].value_counts()
            priority_counts = expired_df['Move_Priority'].value_counts()
            
            # Valor total de inventario que debe moverse
            total_expired_qty = expired_df['QtyOH'].sum()
            avg_expired_qty = expired_df['QtyOH'].mean() if len(expired_df) > 0 else 0
            
            # Top items que deben moverse
            top_expired_items = expired_df.groupby(['ItemNo', 'Description_material']).agg({
                'QtyOH': 'sum',
                'Shelf_Life': 'min',
                'Bin': 'count'
            }).rename(columns={'Bin': 'Locations_Count'}).reset_index()
            top_expired_items = top_expired_items.sort_values('QtyOH', ascending=False)
            
            # Guardar resultados
            self.obsolete_data = merged_df
            self.expired_data = expired_df
            self.top_expired_items = top_expired_items
            
            self.obsolete_metrics = {
                'total_materials_neteable': total_materials_neteable,
                'total_expired': total_expired,
                'expired_percentage': round(expired_percentage, 2),
                'total_expired_qty': int(total_expired_qty) if pd.notna(total_expired_qty) else 0,
                'avg_expired_qty': round(avg_expired_qty, 2) if pd.notna(avg_expired_qty) else 0,
                'unique_expired_items': expired_df['ItemNo'].nunique(),
                'unique_zones': expired_df['Zone'].nunique(),
                'unique_warehouses': expired_df['Wh'].nunique(),
                'locations_with_expired': expired_df['BinID'].nunique(),
                # Conteos por categoría (solo habrá una ya que todos son 0%)
                'critico_count': int(expiry_counts.get('Material Expirado (0%) - Debe moverse de localidad neteable', 0)),
                'muy_expirado_count': 0,  # No aplica
                'expirado_count': 0,      # No aplica  
                'recien_expirado_count': 0, # No aplica
                # Conteos por prioridad de movimiento
                'priority_alta': int(priority_counts.get('Prioridad Alta (>1000 unidades)', 0)),
                'priority_media': int(priority_counts.get('Prioridad Media (100-1000 unidades)', 0)),
                'priority_baja': int(priority_counts.get('Prioridad Baja (<100 unidades)', 0)),
                # Top warehouse y zone
                'top_warehouse': wh_counts.index[0] if len(wh_counts) > 0 else 'N/A',
                'top_warehouse_count': int(wh_counts.iloc[0]) if len(wh_counts) > 0 else 0,
                'top_zone': zone_counts.index[0] if len(zone_counts) > 0 else 'N/A',
                'top_zone_count': int(zone_counts.iloc[0]) if len(zone_counts) > 0 else 0,
            }
            
            print(f"✅ Análisis OBS/EXP NETEABLE completado:")
            print(f"   📊 Total materiales en loc. neteables: {total_materials_neteable:,}")
            print(f"   🚨 Materiales con 0% que deben moverse: {total_expired:,} ({expired_percentage:.2f}%)")
            print(f"   📦 Cantidad total a mover: {int(total_expired_qty):,}")
            print(f"   🏷️ Items únicos a mover: {expired_df['ItemNo'].nunique():,}")
            print(f"   🏢 Warehouses afectados: {expired_df['Wh'].nunique():,}")
            print(f"   📍 Localidades neteables a liberar: {expired_df['BinID'].nunique():,}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error cargando análisis OBS/EXP: {e}")
            import traceback
            traceback.print_exc()
            self._create_sample_obsolete_data()
            return False

    def _create_sample_obsolete_data(self):
        """Crear datos de ejemplo para análisis OBS/EXP"""
        np.random.seed(42)
        n_records = 200
        
        # Datos simulados
        warehouses = ['WH01', 'WH02', 'WH03']
        zones = ['ZONE-A', 'ZONE-B', 'ZONE-C', 'ZONE-D']
        
        data = []
        for i in range(n_records):
            # Simular Shelf_Life: 60% expirados, 40% vigentes
            if np.random.random() < 0.6:
                # Expirados
                shelf_life = np.random.choice([-500, -300, -100, -50, -10, -1], 
                                            p=[0.1, 0.2, 0.3, 0.2, 0.15, 0.05])
            else:
                # Vigentes
                shelf_life = np.random.randint(1, 365)
            
            record = {
                'id': i,
                'Wh': np.random.choice(warehouses),
                'ItemNo': f'ITEM-{1000+i:04d}',
                'Description_material': f'Material Descripción {i}',
                'UM': np.random.choice(['EA', 'LB', 'FT', 'KG']),
                'QtyOH': np.random.randint(1, 2000),
                'Bin': f'BIN-{100+i//5:03d}',
                'Lot': f'LOT-{i:04d}',
                'Shelf_Life': shelf_life,
                'BinID': f'BIN-{100+i//5:03d}',
                'Zone': np.random.choice(zones),
                'Type': np.random.choice(['STORAGE', 'PICK', 'RECEIVING']),
                'Description_location': f'Localidad {i//5}',
                'TagNetable': 'YES'
            }
            data.append(record)
        
        # Crear DataFrames
        all_df = pd.DataFrame(data)
        expired_df = all_df[all_df['Shelf_Life'] <= 0].copy()
        
        # Categorizar
        expired_df['Expiry_Category'] = expired_df['Shelf_Life'].apply(
            lambda x: 'Crítico (< -365 días)' if x < -365
                    else 'Muy Expirado (-180 a -365)' if x < -180
                    else 'Expirado (-30 a -180)' if x < -30
                    else 'Recién Expirado (0 a -30)' if x <= 0
                    else 'Vigente'
        )
        
        expired_df['Qty_Category'] = expired_df['QtyOH'].apply(
            lambda x: 'Alto (>1000)' if x > 1000
                    else 'Medio (100-1000)' if x > 100
                    else 'Bajo (1-100)' if x > 0
                    else 'Sin Stock'
        )
        
        # Top items
        top_items = expired_df.groupby(['ItemNo', 'Description_material']).agg({
            'QtyOH': 'sum',
            'Shelf_Life': 'min',
            'Bin': 'count'
        }).rename(columns={'Bin': 'Locations_Count'}).reset_index()
        top_items = top_items.sort_values('QtyOH', ascending=False)
        
        # Calcular métricas
        total_materials = len(all_df)
        total_expired = len(expired_df)
        expiry_counts = expired_df['Expiry_Category'].value_counts()
        zone_counts = expired_df['Zone'].value_counts()
        wh_counts = expired_df['Wh'].value_counts()
        
        # Guardar datos
        self.obsolete_data = all_df
        self.expired_data = expired_df
        self.top_expired_items = top_items
        
        self.obsolete_metrics = {
            'total_materials_neteable': total_materials,
            'total_expired': total_expired,
            'expired_percentage': round(100 * total_expired / total_materials, 2),
            'total_expired_qty': int(expired_df['QtyOH'].sum()),
            'avg_expired_qty': round(expired_df['QtyOH'].mean(), 2),
            'unique_expired_items': expired_df['ItemNo'].nunique(),
            'unique_zones': expired_df['Zone'].nunique(),
            'unique_warehouses': expired_df['Wh'].nunique(),
            'locations_with_expired': expired_df['BinID'].nunique(),
            'critico_count': int(expiry_counts.get('Crítico (< -365 días)', 0)),
            'muy_expirado_count': int(expiry_counts.get('Muy Expirado (-180 a -365)', 0)),
            'expirado_count': int(expiry_counts.get('Expirado (-30 a -180)', 0)),
            'recien_expirado_count': int(expiry_counts.get('Recién Expirado (0 a -30)', 0)),
            'top_warehouse': wh_counts.index[0] if len(wh_counts) > 0 else 'N/A',
            'top_warehouse_count': int(wh_counts.iloc[0]) if len(wh_counts) > 0 else 0,
            'top_zone': zone_counts.index[0] if len(zone_counts) > 0 else 'N/A',
            'top_zone_count': int(zone_counts.iloc[0]) if len(zone_counts) > 0 else 0,
        }
        
        print(f"🧪 Datos de ejemplo OBS/EXP creados:")
        print(f"   📊 Total materiales en loc. neteables: {total_materials}")
        print(f"   ⚠️ Materiales expirados: {total_expired} ({100*total_expired/total_materials:.2f}%)")

    def create_obsolete_expired_chart(self):
        """Crear gráficos de análisis de materiales obsoletos/expirados"""
        if not hasattr(self, 'expired_data') or self.expired_data is None:
            print("❌ No hay datos expired_data")
            return None
        
        try:
            print("🔄 Iniciando creación de gráfico OBS/EXP...")
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            
            plt.style.use('dark_background')
            
            fig, ((ax1, ax2, ax3), (ax4, ax5, ax6)) = plt.subplots(2, 3, figsize=(30, 16))
            fig.patch.set_facecolor(COLORS['surface'])
            
            print("✅ Configuración inicial completada")
            
            # Verificar que tenemos métricas
            if not hasattr(self, 'obsolete_metrics'):
                print("❌ No hay obsolete_metrics")
                return None
                
            metrics = self.obsolete_metrics
            print(f"✅ Métricas disponibles: {list(metrics.keys())}")
            
            # GRÁFICO 1: Pie Chart - Materiales Expirados vs Vigentes
            print("🔄 Creando gráfico 1: Pie Chart general...")
            try:
                total_materials = metrics.get('total_materials_neteable', 0)
                total_expired = metrics.get('total_expired', 0)
                total_valid = total_materials - total_expired
                
                if total_materials > 0:
                    labels = ['Materiales Vigentes', 'Materiales Expirados']
                    sizes = [total_valid, total_expired]
                    colors = ['#4ECDC4', '#FF6B6B']  # Verde para vigentes, rojo para expirados
                    
                    wedges, texts, autotexts = ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                                    startangle=90, textprops={'color': 'white', 'fontsize': 12})
                    
                    for autotext in autotexts:
                        autotext.set_color('white')
                        autotext.set_fontweight('bold')
                        autotext.set_fontsize(14)
                    
                    ax1.set_title('Distribución General\n(Vigentes vs Expirados)', 
                                color='white', fontsize=16, fontweight='bold', pad=20)
                    
                    # Texto en el centro
                    center_text = f'Total Materiales\n{total_materials:,}'
                    ax1.text(0, 0, center_text, ha='center', va='center', 
                            color='white', fontsize=14, fontweight='bold')
                else:
                    ax1.text(0.5, 0.5, 'No hay datos para mostrar', ha='center', va='center',
                            color='white', fontsize=16, transform=ax1.transAxes)
                
                ax1.set_facecolor('#2C3E50')
                print("✅ Gráfico 1 completado")
                
            except Exception as e1:
                print(f"❌ Error en gráfico 1: {e1}")
                ax1.text(0.5, 0.5, 'Error en Pie Chart', ha='center', va='center',
                        color='red', fontsize=16, transform=ax1.transAxes)
                ax1.set_facecolor('#2C3E50')
            
            # GRÁFICO 2: Barras por categoría de expiración
            print("🔄 Creando gráfico 2: Barras categorías...")
            try:
                categories = ['Crítico\n(< -365)', 'Muy Expirado\n(-180 a -365)', 'Expirado\n(-30 a -180)', 'Recién Expirado\n(0 a -30)']
                values = [
                    metrics.get('critico_count', 0),
                    metrics.get('muy_expirado_count', 0), 
                    metrics.get('expirado_count', 0),
                    metrics.get('recien_expirado_count', 0)
                ]
                colors = ['#8B0000', '#FF4500', '#FF6347', '#FFB366']  # Degradado de rojo
                
                print(f"   Datos categorías: {values}")
                
                bars = ax2.bar(categories, values, color=colors, alpha=0.8, edgecolor='white', linewidth=1)
                
                # Agregar valores en las barras
                for bar, value in zip(bars, values):
                    height = bar.get_height()
                    if height > 0:
                        ax2.text(bar.get_x() + bar.get_width()/2., height + max(values)*0.02,
                            f'{int(value):,}', ha='center', va='bottom', 
                            color='white', fontsize=11, fontweight='bold')
                
                ax2.set_title('Distribución por Severidad de Expiración\n(Cantidad de Materiales)', 
                            color='white', fontsize=16, fontweight='bold', pad=20)
                ax2.set_ylabel('Cantidad de Materiales', color='white', fontsize=14)
                ax2.grid(True, alpha=0.3, color='white', axis='y')
                ax2.set_facecolor('#2C3E50')
                ax2.tick_params(colors='white')
                for spine in ax2.spines.values():
                    spine.set_color('white')
                
                print("✅ Gráfico 2 completado")
                
            except Exception as e2:
                print(f"❌ Error en gráfico 2: {e2}")
                ax2.text(0.5, 0.5, 'Error en Barras Categorías', ha='center', va='center',
                        color='red', fontsize=16, transform=ax2.transAxes)
                ax2.set_facecolor('#2C3E50')
            
            # GRÁFICO 3: Histograma de Shelf_Life
            print("🔄 Creando gráfico 3: Histograma Shelf_Life...")
            try:
                shelf_life_data = self.expired_data['Shelf_Life'].dropna()
                
                if len(shelf_life_data) > 0:
                    # Filtrar valores extremos para mejor visualización
                    filtered_data = shelf_life_data[shelf_life_data >= -1000]  # Limitar a -1000 días
                    
                    n, bins, patches = ax3.hist(filtered_data, bins=50, alpha=0.7, 
                                            color='#FF6B6B', edgecolor='white', linewidth=0.5)
                    
                    # Línea vertical en cero
                    ax3.axvline(x=0, color='yellow', linestyle='--', linewidth=3, alpha=0.9)
                    
                    # Estadísticas
                    mean_val = filtered_data.mean()
                    ax3.axvline(x=mean_val, color='cyan', linestyle='--', 
                            linewidth=2, label=f'Promedio: {mean_val:.0f} días')
                    
                    ax3.set_title('Distribución de Shelf_Life\n(Materiales Expirados)', 
                                color='white', fontsize=16, fontweight='bold', pad=20)
                    ax3.set_xlabel('Shelf_Life (días)', color='white', fontsize=14)
                    ax3.set_ylabel('Cantidad de Materiales', color='white', fontsize=14)
                    ax3.legend(loc='upper left', fontsize=12)
                    ax3.grid(True, alpha=0.3, color='white')
                    ax3.set_facecolor('#2C3E50')
                    ax3.tick_params(colors='white')
                    for spine in ax3.spines.values():
                        spine.set_color('white')
                else:
                    ax3.text(0.5, 0.5, 'No hay datos de Shelf_Life', ha='center', va='center',
                            color='white', fontsize=16, transform=ax3.transAxes)
                    ax3.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 3 completado")
                
            except Exception as e3:
                print(f"❌ Error en gráfico 3: {e3}")
                ax3.text(0.5, 0.5, 'Error en Histograma', ha='center', va='center',
                        color='red', fontsize=16, transform=ax3.transAxes)
                ax3.set_facecolor('#2C3E50')
            
            # GRÁFICO 4: Top 10 Warehouses con materiales expirados
            print("🔄 Creando gráfico 4: Top Warehouses...")
            try:
                wh_counts = self.expired_data['Wh'].value_counts()
                
                if len(wh_counts) > 0:
                    top_wh = wh_counts.head(10)
                    
                    y_pos = range(len(top_wh))
                    bars = ax4.barh(y_pos, top_wh.values, color='#45B7D1', alpha=0.8)
                    
                    # Agregar valores
                    for i, bar in enumerate(bars):
                        width = bar.get_width()
                        ax4.text(width + width*0.01, bar.get_y() + bar.get_height()/2,
                                f'{int(width)}', ha='left', va='center', 
                                color='white', fontsize=10, fontweight='bold')
                    
                    ax4.set_yticks(y_pos)
                    ax4.set_yticklabels(top_wh.index, fontsize=10)
                    ax4.set_title('Top 10 Warehouses\n(Materiales Expirados)', 
                                color='white', fontsize=16, fontweight='bold', pad=20)
                    ax4.set_xlabel('Cantidad de Materiales', color='white', fontsize=14)
                    ax4.grid(True, alpha=0.3, color='white', axis='x')
                    ax4.set_facecolor('#2C3E50')
                    ax4.tick_params(colors='white')
                    for spine in ax4.spines.values():
                        spine.set_color('white')
                else:
                    ax4.text(0.5, 0.5, 'No hay datos de Warehouses', ha='center', va='center',
                            color='white', fontsize=16, transform=ax4.transAxes)
                    ax4.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 4 completado")
                
            except Exception as e4:
                print(f"❌ Error en gráfico 4: {e4}")
                ax4.text(0.5, 0.5, 'Error en Top Warehouses', ha='center', va='center',
                        color='red', fontsize=16, transform=ax4.transAxes)
                ax4.set_facecolor('#2C3E50')
            
            # GRÁFICO 5: Top 10 Zonas con materiales expirados
            print("🔄 Creando gráfico 5: Top Zonas...")
            try:
                zone_counts = self.expired_data['Zone'].value_counts()
                
                if len(zone_counts) > 0:
                    top_zones = zone_counts.head(10)
                    
                    bars = ax5.bar(range(len(top_zones)), top_zones.values, 
                                color='#9B59B6', alpha=0.8, edgecolor='white', linewidth=1)
                    
                    # Agregar valores
                    for bar, value in zip(bars, top_zones.values):
                        height = bar.get_height()
                        ax5.text(bar.get_x() + bar.get_width()/2., height + max(top_zones.values)*0.01,
                            f'{int(value)}', ha='center', va='bottom', 
                            color='white', fontsize=10, fontweight='bold')
                    
                    ax5.set_xticks(range(len(top_zones)))
                    ax5.set_xticklabels(top_zones.index, rotation=45, ha='right', fontsize=10)
                    ax5.set_title('Top 10 Zonas\n(Materiales Expirados)', 
                                color='white', fontsize=16, fontweight='bold', pad=20)
                    ax5.set_ylabel('Cantidad de Materiales', color='white', fontsize=14)
                    ax5.grid(True, alpha=0.3, color='white', axis='y')
                    ax5.set_facecolor('#2C3E50')
                    ax5.tick_params(colors='white')
                    for spine in ax5.spines.values():
                        spine.set_color('white')
                else:
                    ax5.text(0.5, 0.5, 'No hay datos de Zonas', ha='center', va='center',
                            color='white', fontsize=16, transform=ax5.transAxes)
                    ax5.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 5 completado")
                
            except Exception as e5:
                print(f"❌ Error en gráfico 5: {e5}")
                ax5.text(0.5, 0.5, 'Error en Top Zonas', ha='center', va='center',
                        color='red', fontsize=16, transform=ax5.transAxes)
                ax5.set_facecolor('#2C3E50')
            
            # GRÁFICO 6: Scatter plot Qty vs Shelf_Life
            print("🔄 Creando gráfico 6: Scatter Plot...")
            try:
                if len(self.expired_data) > 0:
                    x_data = self.expired_data['Shelf_Life']
                    y_data = self.expired_data['QtyOH']
                    
                    # Filtrar outliers para mejor visualización
                    mask = (x_data >= -1000) & (y_data <= 5000)
                    x_filtered = x_data[mask]
                    y_filtered = y_data[mask]
                    
                    if len(x_filtered) > 0:
                        scatter = ax6.scatter(x_filtered, y_filtered, 
                                            c=x_filtered, cmap='Reds', alpha=0.6, s=20)
                        
                        ax6.set_title('Cantidad vs Shelf_Life\n(Materiales Expirados)', 
                                    color='white', fontsize=16, fontweight='bold', pad=20)
                        ax6.set_xlabel('Shelf_Life (días)', color='white', fontsize=14)
                        ax6.set_ylabel('Cantidad en Stock (QtyOH)', color='white', fontsize=14)
                        ax6.axvline(x=0, color='yellow', linestyle='--', linewidth=2, alpha=0.7)
                        ax6.grid(True, alpha=0.3, color='white')
                        ax6.set_facecolor('#2C3E50')
                        ax6.tick_params(colors='white')
                        for spine in ax6.spines.values():
                            spine.set_color('white')
                        
                        # Colorbar
                        cbar = plt.colorbar(scatter, ax=ax6)
                        cbar.set_label('Shelf_Life (días)', color='white', fontsize=12)
                        cbar.ax.tick_params(colors='white')
                    else:
                        ax6.text(0.5, 0.5, 'No hay datos válidos para scatter', ha='center', va='center',
                                color='white', fontsize=14, transform=ax6.transAxes)
                        ax6.set_facecolor('#2C3E50')
                else:
                    ax6.text(0.5, 0.5, 'No hay datos expirados', ha='center', va='center',
                            color='white', fontsize=16, transform=ax6.transAxes)
                    ax6.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 6 completado")
                
            except Exception as e6:
                print(f"❌ Error en gráfico 6: {e6}")
                ax6.text(0.5, 0.5, 'Error en Scatter Plot', ha='center', va='center',
                        color='red', fontsize=16, transform=ax6.transAxes)
                ax6.set_facecolor('#2C3E50')
            
            plt.tight_layout(pad=3.0)
            
            print("🔄 Convirtiendo a base64...")
            # Convertir a base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', facecolor=COLORS['surface'], 
                    bbox_inches='tight', dpi=150)
            buffer.seek(0)
            chart_base64 = base64.b64encode(buffer.getvalue()).decode()
            plt.close()
            
            print("✅ Gráfico OBS/EXP completado exitosamente")
            return chart_base64
            
        except Exception as e:
            print(f"❌ Error creando gráfico OBS/EXP: {e}")
            import traceback
            print("❌ Traceback completo:")
            traceback.print_exc()
            return None

    def create_fcst_alignment_chart(self):
        """Crear gráfico de alineación FCST"""
        if not hasattr(self, 'fcst_data') or self.fcst_data is None:
            return None
        
        try:
            import matplotlib
            matplotlib.use('Agg')
            
            plt.style.use('dark_background')
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
            fig.patch.set_facecolor(COLORS['surface'])
            
            # GRÁFICO 1: Histograma de diferencias
            diff_data = self.fcst_data['Date_Diff_Days'].dropna()
            
            # Filtrar outliers para mejor visualización
            q1, q3 = diff_data.quantile([0.25, 0.75])
            iqr = q3 - q1
            lower_bound = q1 - 2 * iqr
            upper_bound = q3 + 2 * iqr
            filtered_data = diff_data[(diff_data >= lower_bound) & (diff_data <= upper_bound)]
            
            # Histograma
            n, bins, patches = ax1.hist(filtered_data, bins=50, alpha=0.7, 
                                        density=True, edgecolor='white', linewidth=0.3)
            
            # Colorear barras: verde para alineados, rojo para desalineados
            for i, patch in enumerate(patches):
                if bins[i] < 0:
                    patch.set_facecolor('#FF6B6B')  # Rojo: WO vence antes que FCST
                elif bins[i] == 0:
                    patch.set_facecolor('#4ECDC4')  # Verde: Perfectamente alineado
                else:
                    patch.set_facecolor('#45B7D1')  # Azul: WO vence después que FCST
            
            # Línea vertical en cero
            ax1.axvline(x=0, color='yellow', linestyle='-', linewidth=3, alpha=0.9)
            
            # Estadísticas en el gráfico
            aligned_count = self.fcst_metrics['aligned_count']
            early_count = self.fcst_metrics['early_wo']
            late_count = self.fcst_metrics['late_wo']
            total_count = self.fcst_metrics['total_records']
            
            textstr = f'🟢 Alineados: {aligned_count:,} ({100*aligned_count/total_count:.1f}%)\n🔴 WO Temprano: {early_count:,} ({100*early_count/total_count:.1f}%)\n🔵 WO Tardío: {late_count:,} ({100*late_count/total_count:.1f}%)\nTotal: {total_count:,}'
            props = dict(boxstyle='round', facecolor='black', alpha=0.8)
            ax1.text(0.02, 0.98, textstr, transform=ax1.transAxes, fontsize=11,
                    verticalalignment='top', bbox=props, color='white')
            
            ax1.set_title('Alineación FCST vs WO: DueDt - ReqDate\n(0 = Perfecta Alineación)', 
                        color='white', fontsize=16, fontweight='bold', pad=20)
            ax1.set_xlabel('Diferencia (días)', color='white', fontsize=14)
            ax1.set_ylabel('Densidad', color='white', fontsize=14)
            ax1.grid(True, alpha=0.3, color='white')
            ax1.set_facecolor('#2C3E50')
            ax1.tick_params(colors='white')
            for spine in ax1.spines.values():
                spine.set_color('white')
            
            # GRÁFICO 2: Gráfico de barras por categorías
            categories = ['WO Temprano\n(DueDt < ReqDate)', 'Alineados\n(DueDt = ReqDate)', 'WO Tardío\n(DueDt > ReqDate)']
            values = [early_count, aligned_count, late_count]
            colors = ['#FF6B6B', '#4ECDC4', '#45B7D1']
            
            bars = ax2.bar(categories, values, color=colors, alpha=0.8, edgecolor='white', linewidth=1)
            
            # Agregar valores en las barras
            for bar, value in zip(bars, values):
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height + total_count*0.01,
                        f'{value:,}\n({100*value/total_count:.1f}%)',
                        ha='center', va='bottom', color='white', fontsize=12, fontweight='bold')
            
            ax2.set_title('Distribución de Alineación FCST-WO', 
                        color='white', fontsize=16, fontweight='bold', pad=20)
            ax2.set_ylabel('Cantidad de Registros', color='white', fontsize=14)
            ax2.set_facecolor('#2C3E50')
            ax2.tick_params(colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='y')
            for spine in ax2.spines.values():
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
            print(f"Error creando gráfico FCST: {e}")
            return None

    def create_so_alignment_chart(self):
        """Crear gráfico de alineación SO-WO"""
        if not hasattr(self, 'so_data') or self.so_data is None:
            return None
        
        try:
            import matplotlib
            matplotlib.use('Agg')
            
            plt.style.use('dark_background')
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
            fig.patch.set_facecolor(COLORS['surface'])
            
            # GRÁFICO 1: Histograma de diferencias SO-WO
            diff_data = self.so_data['Date_Diff_Days'].dropna()
            
            # Filtrar outliers para mejor visualización
            q1, q3 = diff_data.quantile([0.25, 0.75])
            iqr = q3 - q1
            lower_bound = q1 - 2 * iqr
            upper_bound = q3 + 2 * iqr
            filtered_data = diff_data[(diff_data >= lower_bound) & (diff_data <= upper_bound)]
            
            # Histograma
            n, bins, patches = ax1.hist(filtered_data, bins=50, alpha=0.7, 
                                        density=True, edgecolor='white', linewidth=0.3)
            
            # Colorear barras: rojo para WO temprana, verde para alineados, azul para WO tardía
            for i, patch in enumerate(patches):
                if bins[i] < 0:
                    patch.set_facecolor('#FF6B6B')  # Rojo: WO vence antes que SO
                elif bins[i] == 0:
                    patch.set_facecolor('#4ECDC4')  # Verde: Perfectamente alineado
                else:
                    patch.set_facecolor('#45B7D1')  # Azul: WO vence después que SO
            
            # Línea vertical en cero
            ax1.axvline(x=0, color='yellow', linestyle='-', linewidth=3, alpha=0.9)
            
            # Estadísticas en el gráfico
            aligned_count = self.so_metrics['aligned_count']
            early_count = self.so_metrics['early_wo']
            late_count = self.so_metrics['late_wo']
            total_count = self.so_metrics['total_records']
            
            textstr = f'🟢 Alineados: {aligned_count:,} ({100*aligned_count/total_count:.1f}%)\n🔴 WO Temprana: {early_count:,} ({100*early_count/total_count:.1f}%)\n🔵 WO Tardía: {late_count:,} ({100*late_count/total_count:.1f}%)\nTotal: {total_count:,}'
            props = dict(boxstyle='round', facecolor='black', alpha=0.8)
            ax1.text(0.02, 0.98, textstr, transform=ax1.transAxes, fontsize=11,
                    verticalalignment='top', bbox=props, color='white')
            
            ax1.set_title('Alineación SO vs WO: DueDt - Prd_Dt\n(0 = Perfecta Alineación)', 
                        color='white', fontsize=16, fontweight='bold', pad=20)
            ax1.set_xlabel('Diferencia (días)', color='white', fontsize=14)
            ax1.set_ylabel('Densidad', color='white', fontsize=14)
            ax1.grid(True, alpha=0.3, color='white')
            ax1.set_facecolor('#2C3E50')
            ax1.tick_params(colors='white')
            for spine in ax1.spines.values():
                spine.set_color('white')
            
            # GRÁFICO 2: Gráfico de barras por categorías
            categories = ['WO Temprana\n(DueDt < Prd_Dt)', 'Alineados\n(DueDt = Prd_Dt)', 'WO Tardía\n(DueDt > Prd_Dt)']
            values = [early_count, aligned_count, late_count]
            colors = ['#FF6B6B', '#4ECDC4', '#45B7D1']
            
            bars = ax2.bar(categories, values, color=colors, alpha=0.8, edgecolor='white', linewidth=1)
            
            # Agregar valores en las barras
            for bar, value in zip(bars, values):
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height + total_count*0.01,
                        f'{value:,}\n({100*value/total_count:.1f}%)',
                        ha='center', va='bottom', color='white', fontsize=12, fontweight='bold')
            
            ax2.set_title('Distribución de Alineación SO-WO', 
                        color='white', fontsize=16, fontweight='bold', pad=20)
            ax2.set_ylabel('Cantidad de Registros', color='white', fontsize=14)
            ax2.set_facecolor('#2C3E50')
            ax2.tick_params(colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='y')
            for spine in ax2.spines.values():
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
            print(f"Error creando gráfico SO-WO: {e}")
            return None

    def create_wo_materials_chart(self):
        """Crear gráfico de análisis WO con/sin materiales"""
        if not hasattr(self, 'wo_materials_data') or self.wo_materials_data is None:
            print("❌ No hay datos wo_materials_data")
            return None
        
        try:
            print("🔄 Iniciando creación de gráfico WO Materials...")
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            
            plt.style.use('dark_background')
            
            fig, ((ax1, ax2, ax3), (ax4, ax5, ax6)) = plt.subplots(2, 3, figsize=(30, 16))
            fig.patch.set_facecolor(COLORS['surface'])
            
            print("✅ Configuración inicial completada")
            
            # Verificar que tenemos métricas
            if not hasattr(self, 'wo_materials_metrics'):
                print("❌ No hay wo_materials_metrics")
                return None
                
            metrics = self.wo_materials_metrics
            print(f"✅ Métricas disponibles: {list(metrics.keys())}")
            
            # GRÁFICO 1: Pie Chart - WO Con/Sin Materiales
            print("🔄 Creando gráfico 1: Pie Chart...")
            try:
                labels = ['Con Materiales', 'Sin Materiales']
                sizes = [metrics.get('wo_with_materials', 0), metrics.get('wo_without_materials', 0)]
                colors = ['#4ECDC4', '#FF6B6B']  # Verde para con materiales, rojo para sin materiales
                
                print(f"   Datos pie: {sizes}")
                
                if sum(sizes) > 0:
                    wedges, texts, autotexts = ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                                    startangle=90, textprops={'color': 'white', 'fontsize': 12})
                    
                    # Mejorar el texto del pie chart
                    for autotext in autotexts:
                        autotext.set_color('white')
                        autotext.set_fontweight('bold')
                        autotext.set_fontsize(14)
                    
                    ax1.set_title('Distribución General de WO\n(Con/Sin Materiales)', 
                                color='white', fontsize=16, fontweight='bold', pad=20)
                    
                    # Agregar estadísticas en el pie chart
                    center_text = f'Total WO\n{metrics.get("total_wo", 0):,}'
                    ax1.text(0, 0, center_text, ha='center', va='center', 
                            color='white', fontsize=14, fontweight='bold')
                else:
                    ax1.text(0.5, 0.5, 'No hay datos para mostrar', ha='center', va='center',
                            color='white', fontsize=16, transform=ax1.transAxes)
                    ax1.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 1 completado")
                
            except Exception as e1:
                print(f"❌ Error en gráfico 1: {e1}")
                ax1.text(0.5, 0.5, 'Error en Pie Chart', ha='center', va='center',
                        color='red', fontsize=16, transform=ax1.transAxes)
                ax1.set_facecolor('#2C3E50')
            
            # GRÁFICO 2: Barras por tipo Srt
            print("🔄 Creando gráfico 2: Barras Srt...")
            try:
                srt_categories = ['GF', 'GFA', 'KZ']
                srt_totals = [metrics.get('srt_gf_total', 0), metrics.get('srt_gfa_total', 0), metrics.get('srt_kz_total', 0)]
                srt_without = [metrics.get('srt_gf_without', 0), metrics.get('srt_gfa_without', 0), metrics.get('srt_kz_without', 0)]
                srt_with = [total - without for total, without in zip(srt_totals, srt_without)]
                
                print(f"   Datos Srt totals: {srt_totals}")
                print(f"   Datos Srt without: {srt_without}")
                print(f"   Datos Srt with: {srt_with}")
                
                x = range(len(srt_categories))
                width = 0.35
                
                bars1 = ax2.bar([i - width/2 for i in x], srt_with, width, 
                            label='Con Materiales', color='#4ECDC4', alpha=0.8)
                bars2 = ax2.bar([i + width/2 for i in x], srt_without, width,
                            label='Sin Materiales', color='#FF6B6B', alpha=0.8)
                
                # Agregar valores en las barras
                for bars in [bars1, bars2]:
                    for bar in bars:
                        height = bar.get_height()
                        if height > 0:
                            ax2.text(bar.get_x() + bar.get_width()/2., height + max(srt_totals)*0.01,
                                f'{int(height)}', ha='center', va='bottom', 
                                color='white', fontsize=11, fontweight='bold')
                
                ax2.set_title('Distribución por Tipo Srt\n(GF, GFA, KZ)', 
                            color='white', fontsize=16, fontweight='bold', pad=20)
                ax2.set_xlabel('Tipo Srt', color='white', fontsize=14)
                ax2.set_ylabel('Cantidad de WO', color='white', fontsize=14)
                ax2.set_xticks(x)
                ax2.set_xticklabels(srt_categories)
                ax2.legend(loc='upper right', fontsize=12)
                ax2.grid(True, alpha=0.3, color='white', axis='y')
                ax2.set_facecolor('#2C3E50')
                ax2.tick_params(colors='white')
                for spine in ax2.spines.values():
                    spine.set_color('white')
                
                print("✅ Gráfico 2 completado")
                
            except Exception as e2:
                print(f"❌ Error en gráfico 2: {e2}")
                ax2.text(0.5, 0.5, 'Error en Barras Srt', ha='center', va='center',
                        color='red', fontsize=16, transform=ax2.transAxes)
                ax2.set_facecolor('#2C3E50')
            
            # GRÁFICO 3: Histograma de cantidad de materiales por WO
            print("🔄 Creando gráfico 3: Histograma...")
            try:
                wo_with_materials = self.wo_materials_data[self.wo_materials_data['Has_Materials']]
                print(f"   WO con materiales: {len(wo_with_materials)}")
                
                if len(wo_with_materials) > 0:
                    materials_counts = wo_with_materials['Materials_Count']
                    print(f"   Materials_counts stats: min={materials_counts.min()}, max={materials_counts.max()}, mean={materials_counts.mean():.2f}")
                    
                    max_count = int(materials_counts.max())
                    bins = min(20, max_count) if max_count > 0 else 10
                    
                    n, bins_array, patches = ax3.hist(materials_counts, bins=bins, 
                                            alpha=0.7, color='#45B7D1', edgecolor='white', linewidth=0.5)
                    
                    # Colorear barras gradualmente
                    for i, patch in enumerate(patches):
                        patch.set_facecolor('#45B7D1')
                    
                    mean_val = materials_counts.mean()
                    ax3.axvline(x=mean_val, color='yellow', linestyle='--', 
                            linewidth=2, label=f'Promedio: {mean_val:.1f}')
                    
                    ax3.set_title('Distribución: Cantidad de Materiales por WO\n(Solo WO con Materiales)', 
                                color='white', fontsize=16, fontweight='bold', pad=20)
                    ax3.set_xlabel('Cantidad de Materiales', color='white', fontsize=14)
                    ax3.set_ylabel('Cantidad de WO', color='white', fontsize=14)
                    ax3.legend(loc='upper right', fontsize=12)
                    ax3.grid(True, alpha=0.3, color='white')
                    ax3.set_facecolor('#2C3E50')
                    ax3.tick_params(colors='white')
                    for spine in ax3.spines.values():
                        spine.set_color('white')
                else:
                    ax3.text(0.5, 0.5, 'No hay WO con materiales', ha='center', va='center',
                            color='white', fontsize=16, transform=ax3.transAxes)
                    ax3.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 3 completado")
                
            except Exception as e3:
                print(f"❌ Error en gráfico 3: {e3}")
                ax3.text(0.5, 0.5, 'Error en Histograma', ha='center', va='center',
                        color='red', fontsize=16, transform=ax3.transAxes)
                ax3.set_facecolor('#2C3E50')
            
            # GRÁFICO 4: Top 10 WO sin materiales (por valor OpnQ)
            print("🔄 Creando gráfico 4: Top 10...")
            try:
                wo_without = self.wo_materials_data[~self.wo_materials_data['Has_Materials']]
                print(f"   WO sin materiales: {len(wo_without)}")
                
                if len(wo_without) > 0:
                    # Convertir OpnQ a numérico
                    wo_without = wo_without.copy()
                    wo_without['OpnQ'] = pd.to_numeric(wo_without['OpnQ'], errors='coerce')
                    wo_without = wo_without.dropna(subset=['OpnQ'])
                    
                    if len(wo_without) > 0:
                        top_wo_without = wo_without.nlargest(10, 'OpnQ')[['WONo', 'OpnQ', 'Srt']]
                        print(f"   Top 10 WO sin materiales: {len(top_wo_without)}")
                        
                        if len(top_wo_without) > 0:
                            y_pos = range(len(top_wo_without))
                            colors_map = {'GF': '#FF6B6B', 'GFA': '#FFB366', 'KZ': '#66B2FF'}
                            bar_colors = [colors_map.get(str(srt), '#888888') for srt in top_wo_without['Srt']]
                            
                            bars = ax4.barh(y_pos, top_wo_without['OpnQ'], color=bar_colors, alpha=0.8)
                            
                            # Agregar valores al final de las barras
                            for i, bar in enumerate(bars):
                                width = bar.get_width()
                                ax4.text(width + width*0.01, bar.get_y() + bar.get_height()/2,
                                        f'{int(width)}', ha='left', va='center', 
                                        color='white', fontsize=10, fontweight='bold')
                            
                            # Personalizar etiquetas del eje Y
                            wo_labels = [f"{str(row['WONo'])[-8:]}({str(row['Srt'])})" for _, row in top_wo_without.iterrows()]
                            ax4.set_yticks(y_pos)
                            ax4.set_yticklabels(wo_labels, fontsize=10)
                            
                            ax4.set_title('Top 10 WO Sin Materiales\n(Por Cantidad Abierta OpnQ)', 
                                        color='white', fontsize=16, fontweight='bold', pad=20)
                            ax4.set_xlabel('Cantidad Abierta (OpnQ)', color='white', fontsize=14)
                            ax4.grid(True, alpha=0.3, color='white', axis='x')
                            ax4.set_facecolor('#2C3E50')
                            ax4.tick_params(colors='white')
                            for spine in ax4.spines.values():
                                spine.set_color('white')
                                
                            # Leyenda de colores Srt
                            legend_elements = [plt.Rectangle((0,0),1,1, facecolor=colors_map['GF'], label='GF'),
                                            plt.Rectangle((0,0),1,1, facecolor=colors_map['GFA'], label='GFA'),
                                            plt.Rectangle((0,0),1,1, facecolor=colors_map['KZ'], label='KZ')]
                            ax4.legend(handles=legend_elements, loc='lower right', fontsize=10)
                        else:
                            ax4.text(0.5, 0.5, 'No hay datos válidos para Top 10', ha='center', va='center',
                                    color='white', fontsize=14, transform=ax4.transAxes)
                            ax4.set_facecolor('#2C3E50')
                    else:
                        ax4.text(0.5, 0.5, 'No hay OpnQ válidas', ha='center', va='center',
                                color='white', fontsize=14, transform=ax4.transAxes)
                        ax4.set_facecolor('#2C3E50')
                else:
                    ax4.text(0.5, 0.5, 'No hay WO sin materiales', ha='center', va='center',
                            color='white', fontsize=16, transform=ax4.transAxes)
                    ax4.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 4 completado")
                
            except Exception as e4:
                print(f"❌ Error en gráfico 4: {e4}")
                ax4.text(0.5, 0.5, 'Error en Top 10', ha='center', va='center',
                        color='red', fontsize=16, transform=ax4.transAxes)
                ax4.set_facecolor('#2C3E50')
            
            # GRÁFICO 5: Barras por WO_Status
            print("🔄 Creando gráfico 5: Barras Status...")
            try:
                # Obtener datos de status dinámicamente
                status_data = []
                for key in metrics.keys():
                    if key.startswith('status_') and key.endswith('_total'):
                        status = key.replace('status_', '').replace('_total', '')
                        total = metrics.get(f'status_{status}_total', 0)
                        without = metrics.get(f'status_{status}_without', 0)
                        with_mat = total - without
                        status_data.append({'status': status, 'total': total, 'without': without, 'with': with_mat})
                
                if status_data:
                    status_df = pd.DataFrame(status_data).sort_values('total', ascending=False)
                    
                    x = range(len(status_df))
                    width = 0.35
                    
                    bars1 = ax5.bar([i - width/2 for i in x], status_df['with'], width, 
                                label='Con Materiales', color='#4ECDC4', alpha=0.8)
                    bars2 = ax5.bar([i + width/2 for i in x], status_df['without'], width,
                                label='Sin Materiales', color='#FF6B6B', alpha=0.8)
                    
                    # Agregar valores en las barras
                    for bars in [bars1, bars2]:
                        for bar in bars:
                            height = bar.get_height()
                            if height > 0:
                                ax5.text(bar.get_x() + bar.get_width()/2., height + max(status_df['total'])*0.01,
                                    f'{int(height)}', ha='center', va='bottom', 
                                    color='white', fontsize=9, fontweight='bold')
                    
                    ax5.set_title('Distribución por WO_Status\n(Todos los Status)', 
                                color='white', fontsize=14, fontweight='bold', pad=15)
                    ax5.set_xlabel('WO Status', color='white', fontsize=12)
                    ax5.set_ylabel('Cantidad de WO', color='white', fontsize=12)
                    ax5.set_xticks(x)
                    ax5.set_xticklabels(status_df['status'])
                    ax5.legend(loc='upper right', fontsize=10)
                    ax5.grid(True, alpha=0.3, color='white', axis='y')
                    ax5.set_facecolor('#2C3E50')
                    ax5.tick_params(colors='white')
                    for spine in ax5.spines.values():
                        spine.set_color('white')
                else:
                    ax5.text(0.5, 0.5, 'No hay datos de Status', ha='center', va='center',
                            color='white', fontsize=14, transform=ax5.transAxes)
                    ax5.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 5 completado")
                
            except Exception as e5:
                print(f"❌ Error en gráfico 5: {e5}")
                ax5.text(0.5, 0.5, 'Error en Status', ha='center', va='center',
                        color='red', fontsize=14, transform=ax5.transAxes)
                ax5.set_facecolor('#2C3E50')

            # GRÁFICO 6: Pie Chart solo de WO sin materiales por Status
            print("🔄 Creando gráfico 6: Pie Status sin materiales...")
            try:
                # Usar status_data del gráfico anterior
                if 'status_data' in locals():
                    without_data = [(data['status'], data['without']) for data in status_data if data['without'] > 0]
                    
                    if without_data and sum(item[1] for item in without_data) > 0:
                        labels = [item[0] for item in without_data]
                        sizes = [item[1] for item in without_data]
                        colors = ['#FF6B6B', '#FFB366', '#66B2FF', '#9B59B6', '#2ECC71', '#F39C12'][:len(labels)]
                        
                        wedges, texts, autotexts = ax6.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                                        startangle=90, textprops={'color': 'white', 'fontsize': 10})
                        
                        for autotext in autotexts:
                            autotext.set_color('white')
                            autotext.set_fontweight('bold')
                            autotext.set_fontsize(11)
                        
                        ax6.set_title('WO Sin Materiales\npor Status', 
                                    color='white', fontsize=14, fontweight='bold', pad=15)
                    else:
                        ax6.text(0.5, 0.5, 'Todas las WO\ntienen materiales', ha='center', va='center',
                                color='white', fontsize=14, transform=ax6.transAxes)
                        ax6.set_facecolor('#2C3E50')
                else:
                    ax6.text(0.5, 0.5, 'No hay datos de Status', ha='center', va='center',
                            color='white', fontsize=14, transform=ax6.transAxes)
                    ax6.set_facecolor('#2C3E50')
                
                print("✅ Gráfico 6 completado")
                
            except Exception as e6:
                print(f"❌ Error en gráfico 6: {e6}")
                ax6.text(0.5, 0.5, 'Error en Pie Status', ha='center', va='center',
                        color='red', fontsize=14, transform=ax6.transAxes)
                ax6.set_facecolor('#2C3E50')
            
            plt.tight_layout(pad=3.0)
            
            print("🔄 Convirtiendo a base64...")
            # Convertir a base64
            buffer = io.BytesIO()
            plt.savefig(buffer, format='png', facecolor=COLORS['surface'], 
                    bbox_inches='tight', dpi=150)
            buffer.seek(0)
            chart_base64 = base64.b64encode(buffer.getvalue()).decode()
            plt.close()
            
            print("✅ Gráfico WO Materials completado exitosamente")
            return chart_base64
            
        except Exception as e:
            print(f"❌ Error creando gráfico WO-Materials: {e}")
            import traceback
            print("❌ Traceback completo:")
            traceback.print_exc()
            return None


def create_top_items_table(analyzer):
    """Crear tabla de top items expirados"""
    try:
        if not hasattr(analyzer, 'top_expired_items') or analyzer.top_expired_items is None:
            return ft.Text("No hay datos de top items", color=COLORS['text_secondary'])
        
        top_10 = analyzer.top_expired_items.head(10)
        
        if len(top_10) == 0:
            return ft.Text("No hay items expirados", color=COLORS['text_secondary'])
        
        # Crear filas de la tabla
        rows = []
        for idx, row in top_10.iterrows():
            table_row = ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(str(row['ItemNo'])[:20], size=10, color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(str(row['Description_material'])[:30], size=10, color=COLORS['text_secondary'])),
                    ft.DataCell(ft.Text(f"{int(row['QtyOH']):,}", size=10, color=COLORS['accent'])),
                    ft.DataCell(ft.Text(f"{int(row['Shelf_Life'])}", size=10, color=COLORS['error'])),
                    ft.DataCell(ft.Text(f"{int(row['Locations_Count'])}", size=10, color=COLORS['warning']))
                ]
            )
            rows.append(table_row)
        
        # Crear tabla
        table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Item No", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("Descripción", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("Qty Total", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("Shelf_Life", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("# Localidades", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD))
            ],
            rows=rows,
            heading_row_color=COLORS['secondary'],
            data_row_color=COLORS['primary'],
            border_radius=8
        )
        
        return ft.Container(
            content=table,
            padding=ft.padding.all(10),
            bgcolor=COLORS['primary'],
            border_radius=8
        )
        
    except Exception as e:
        print(f"Error creando tabla top items: {e}")
        return ft.Text(f"Error: {str(e)}", color=COLORS['error'])

def create_metric_card(title, value, subtitle="", color=COLORS['accent']):
    """Crear tarjeta de métrica"""
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

def create_gaussian_cards(analyzer):
    """Crear tarjetas para cada campana CORREGIDAS"""
    if analyzer.df is None or len(analyzer.df) == 0:
        return ft.Row([])
    
    # CORREGIDO: Usar las variables correctas
    bs_data = analyzer.df["Diff_Ship_Req"].dropna()        # ShipDate - ReqDate
    bs_problems = len(bs_data[bs_data < 0])                 # Material llega DESPUÉS (PROBLEMAS)
    bs_good = len(bs_data[bs_data > 0])                     # Material llega ANTES (BUENO - BackSchedule correcto)
    bs_ontime = len(bs_data[bs_data == 0])                  # Material llega JUSTO A TIEMPO
    bs_total = len(bs_data)
    bs_leadtime_avg = bs_data[bs_data > 0].mean() if bs_good > 0 else 0
    
    """aqui se comenzo a modificar """
    # NUEVO: Calcular el mayor BackSchedule (valor máximo positivo)
    bs_max_backschedule = bs_data[bs_data > 0].max() if bs_good > 0 else 0
    

    # NUEVO: Análisis de Material_Type para BackSchedule Justo a Tiempo
    bs_ontime_data = analyzer.df[analyzer.df["Diff_Ship_Req"] == 0]
    bs_ontime_invoiceable = len(bs_ontime_data[bs_ontime_data["Material_Type"] == "Invoiceable"])
    bs_ontime_raw = len(bs_ontime_data[bs_ontime_data["Material_Type"] == "Raw Material"])



    orig_data = analyzer.df["Diff_Ship_Original"].dropna()  # ShipDate - ReqDate_Original
    orig_problems = len(orig_data[orig_data < 0])           # Material llega DESPUÉS (PROBLEMAS)
    orig_good = len(orig_data[orig_data > 0])               # Material llega ANTES (BUENO - BackSchedule correcto)
    orig_ontime = len(orig_data[orig_data == 0])            # Material llega JUSTO A TIEMPO
    orig_total = len(orig_data)
    orig_leadtime_avg = orig_data[orig_data > 0].mean() if orig_good > 0 else 0
    
    # NUEVO: Calcular el mayor BackSchedule original
    orig_max_backschedule = orig_data[orig_data > 0].max() if orig_good > 0 else 0
    
    # NUEVO: Análisis de Material_Type para Original Justo a Tiempo
    orig_ontime_data = analyzer.df[analyzer.df["Diff_Ship_Original"] == 0]
    orig_ontime_invoiceable = len(orig_ontime_data[orig_ontime_data["Material_Type"] == "Invoiceable"])
    orig_ontime_raw = len(orig_ontime_data[orig_ontime_data["Material_Type"] == "Raw Material"])







    bs_card = ft.Card(
        content=ft.Container(
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12,
            content=ft.Column([
                ft.Row([
                    ft.Text("🔵", size=24),
                    ft.Text("BackSchedule", size=18, color=COLORS['accent'], weight=ft.FontWeight.BOLD)
                ]),
                ft.Text("ShipDate - ReqDate", size=12, color=COLORS['text_secondary']),
                ft.Divider(color=COLORS['secondary']),
                ft.Row([
                    ft.Column([
                        ft.Text("❌ Problemas", size=12, color=COLORS['text_secondary']),
                        ft.Text(f"{bs_problems:,}", size=18, color=COLORS['error'], weight=ft.FontWeight.BOLD),
                        ft.Text(f"{100*bs_problems/bs_total:.1f}%", size=10, color=COLORS['text_secondary'])
                    ], expand=1),
                    ft.Column([
                        ft.Text("✅ BackSchedule", size=12, color=COLORS['text_secondary']),
                        ft.Text(f"{bs_good:,}", size=18, color=COLORS['success'], weight=ft.FontWeight.BOLD),
                        ft.Text(f"{100*bs_good/bs_total:.1f}%", size=10, color=COLORS['text_secondary'])
                    ], expand=1)
                ]),
                ft.Row([
                ft.Column([
                    ft.Text("⭕ Justo a tiempo", size=12, color=COLORS['text_secondary']),
                    ft.Text(f"{bs_ontime:,}", size=18, color=COLORS['warning'], weight=ft.FontWeight.BOLD),
                    ft.Text(f"{100*bs_ontime/bs_total:.1f}%", size=10, color=COLORS['text_secondary']),
                    ft.Row([
                        ft.Column([
                            ft.Text("📦 Invoiceable", size=9, color=COLORS['text_secondary']),
                            ft.Text(f"{bs_ontime_invoiceable:,}", size=11, color=COLORS['accent'])
                        ], expand=1),
                        ft.Column([
                            ft.Text("🔧 Raw Material", size=9, color=COLORS['text_secondary']),
                            ft.Text(f"{bs_ontime_raw:,}", size=11, color=COLORS['success'])
                        ], expand=1)
                    ])
                ], expand=2)
            ]),
                ft.Divider(color=COLORS['secondary']),
                ft.Text(f"BackSchedule promedio: {bs_leadtime_avg:.1f} días", size=12, color=COLORS['success'])
            ], spacing=10)
        ),
        elevation=8
    )
    
    orig_card = ft.Card(
        content=ft.Container(
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12,
            content=ft.Column([
                ft.Row([
                    ft.Text("🟢", size=24),
                    ft.Text("Sistema Original", size=18, color=COLORS['success'], weight=ft.FontWeight.BOLD)
                ]),
                ft.Text("ShipDate - ReqDate_Original", size=12, color=COLORS['text_secondary']),
                ft.Divider(color=COLORS['secondary']),
                ft.Row([
                    ft.Column([
                        ft.Text("❌ Problemas", size=12, color=COLORS['text_secondary']),
                        ft.Text(f"{orig_problems:,}", size=18, color=COLORS['error'], weight=ft.FontWeight.BOLD),
                        ft.Text(f"{100*orig_problems/orig_total:.1f}%", size=10, color=COLORS['text_secondary'])
                    ], expand=1),
                    ft.Column([
                        ft.Text("✅ BackSchedule", size=12, color=COLORS['text_secondary']),
                        ft.Text(f"{orig_good:,}", size=18, color=COLORS['success'], weight=ft.FontWeight.BOLD),
                        ft.Text(f"{100*orig_good/orig_total:.1f}%", size=10, color=COLORS['text_secondary'])
                    ], expand=1)
                ]),
                ft.Row([
                    ft.Column([
                        ft.Text("⭕ Justo a tiempo", size=12, color=COLORS['text_secondary']),
                        ft.Text(f"{orig_ontime:,}", size=18, color=COLORS['warning'], weight=ft.FontWeight.BOLD),
                        ft.Text(f"{100*orig_ontime/orig_total:.1f}%", size=10, color=COLORS['text_secondary']),
                        ft.Row([
                            ft.Column([
                                ft.Text("📦 Invoiceable", size=9, color=COLORS['text_secondary']),
                                ft.Text(f"{orig_ontime_invoiceable:,}", size=11, color=COLORS['accent'])
                            ], expand=1),
                            ft.Column([
                                ft.Text("🔧 Raw Material", size=9, color=COLORS['text_secondary']),
                                ft.Text(f"{orig_ontime_raw:,}", size=11, color=COLORS['success'])
                            ], expand=1)
                        ])
                    ], expand=2)
                ]),
                ft.Divider(color=COLORS['secondary']),
                ft.Text(f"BackSchedule promedio: {orig_leadtime_avg:.1f} días", size=12, color=COLORS['success'])
            ], spacing=10)
        ),
        elevation=8
    )

    # NUEVAS TARJETAS: Mayor BackSchedule
    bs_max_card = ft.Card(
        content=ft.Container(
            padding=ft.padding.all(20),
            bgcolor=COLORS['secondary'],
            border_radius=12,
            content=ft.Column([
                ft.Row([
                    ft.Text("🚀", size=24),
                    ft.Text("Mayor BackSchedule", size=16, color=COLORS['accent'], weight=ft.FontWeight.BOLD)
                ]),
                ft.Text("BackSchedule Sistema", size=12, color=COLORS['text_secondary']),
                ft.Divider(color=COLORS['secondary']),
                ft.Text(f"{bs_max_backschedule:.0f}", size=36, color=COLORS['accent'], weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER),
                ft.Text("días antes", size=14, color=COLORS['text_secondary'], text_align=ft.TextAlign.CENTER),
                ft.Container(
                    content=ft.Text("🔵 BackSchedule", size=12, color=COLORS['accent'], weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER),
                    bgcolor=COLORS['primary'],
                    padding=ft.padding.all(8),
                    border_radius=6
                )
            ], spacing=8)
        ),
        elevation=8
    )

    orig_max_card = ft.Card(
        content=ft.Container(
            padding=ft.padding.all(20),
            bgcolor=COLORS['secondary'],
            border_radius=12,
            content=ft.Column([
                ft.Row([
                    ft.Text("🚀", size=24),
                    ft.Text("Mayor BackSchedule", size=16, color=COLORS['success'], weight=ft.FontWeight.BOLD)
                ]),
                ft.Text("Sistema Original", size=12, color=COLORS['text_secondary']),
                ft.Divider(color=COLORS['secondary']),
                ft.Text(f"{orig_max_backschedule:.0f}", size=36, color=COLORS['success'], weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER),
                ft.Text("días antes", size=14, color=COLORS['text_secondary'], text_align=ft.TextAlign.CENTER),
                ft.Container(
                    content=ft.Text("🟢 Original", size=12, color=COLORS['success'], weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER),
                    bgcolor=COLORS['primary'],
                    padding=ft.padding.all(8),
                    border_radius=6
                )
            ], spacing=8)
        ),
        elevation=8
    )

    # Organizar en dos filas
    return ft.Row([bs_card, orig_card, bs_max_card, orig_max_card], spacing=20)

def create_export_card(analyzer, export_callback):
    """Crear tarjeta de exportación CORREGIDA con callback"""
    summary = analyzer.get_effectiveness_summary()
    
    # Estado local para el botón
    export_button_ref = ft.Ref[ft.ElevatedButton]()
    status_text_ref = ft.Ref[ft.Text]()
    
    def handle_export(_):
        # Cambiar estado del botón a "procesando"
        export_button_ref.current.text = "⏳ Exportando..."
        export_button_ref.current.disabled = True
        status_text_ref.current.value = "🔄 Procesando exportación..."
        status_text_ref.current.color = COLORS['warning']
        export_button_ref.current.update()
        status_text_ref.current.update()
        
        # Llamar la función de exportación
        try:
            result = analyzer.export_effectiveness_analysis_to_excel()
            
            if result["success"]:
                status_text_ref.current.value = f"✅ {result['message']} - Abriendo archivo..."
                status_text_ref.current.color = COLORS['success']
                export_button_ref.current.text = "✅ Completado"
                
                # Abrir el archivo automáticamente
                try:
                    import subprocess
                    import platform
                    
                    file_path = result['path']
                    
                    if platform.system() == 'Darwin':       # macOS
                        subprocess.call(('open', file_path))
                    elif platform.system() == 'Windows':   # Windows
                        subprocess.call(('start', file_path), shell=True)
                    else:                                   # Linux
                        subprocess.call(('xdg-open', file_path))
                        
                    status_text_ref.current.value = f"✅ Exportado y abierto: {result['message']}"
                    
                except Exception as open_error:
                    status_text_ref.current.value = f"✅ Exportado: {result['message']} (No se pudo abrir automáticamente)"
                    print(f"Error abriendo archivo: {open_error}")
                    
            else:
                status_text_ref.current.value = f"❌ {result['message']}"
                status_text_ref.current.color = COLORS['error']
                export_button_ref.current.text = "❌ Error"



        except Exception as e:
            status_text_ref.current.value = f"❌ Error: {str(e)}"
            status_text_ref.current.color = COLORS['error']
            export_button_ref.current.text = "❌ Error"
        
        # Actualizar UI
        status_text_ref.current.update()
        export_button_ref.current.update()
        
        # Restaurar botón después de 3 segundos
        import time
        import threading
        def restore_button():
            time.sleep(3)
            export_button_ref.current.text = "📊 Exportar"
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
                    ft.Text("📤 Exportar Análisis Completo", 
                            size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                    ft.ElevatedButton(
                        "📊 Exportar",
                        on_click=handle_export,
                        bgcolor=COLORS['error'],
                        color=COLORS['text_primary'],
                        ref=export_button_ref
                    )
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                
                # Área de estado
                ft.Text(
                    "Listo para exportar",
                    size=12,
                    color=COLORS['text_secondary'],
                    ref=status_text_ref
                ),
                
                ft.Divider(color=COLORS['secondary']),
                ft.Text("📋 Se exportarán las siguientes hojas:", size=12, color=COLORS['text_secondary']),
                ft.Column([
                    ft.Text("• BS_Problemas, BS_JustoATiempo, BS_Buenos", size=11, color=COLORS['accent']),
                    ft.Text("• Orig_Problemas, Orig_JustoATiempo, Orig_Buenos", size=11, color=COLORS['success']),
                    
                ], spacing=2),
                ft.Divider(color=COLORS['secondary']),
                ft.Row([
                    ft.Column([
                        ft.Text("🔵 BackSchedule", size=14, color=COLORS['accent']),
                        ft.Text(f"{summary.get('backschedule_problems', 0):,} problemas", 
                                size=16, color=COLORS['error'])
                    ], expand=1),
                    ft.Column([
                        ft.Text("🟢 Original", size=14, color=COLORS['success']),
                        ft.Text(f"{summary.get('original_problems', 0):,} problemas", 
                                size=16, color=COLORS['error'])
                    ], expand=1)
                ]),
                ft.Text(f"🏆 Mejor: {summary.get('better_system', 'N/A')}", 
                        size=14, color=COLORS['warning'], weight=ft.FontWeight.BOLD)
            ], spacing=10)
        ),
        elevation=8
    )

def main(page: ft.Page):
    page.title = "BackSchedule Analytics Dashboard - INTERPRETACIÓN FINAL"
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = COLORS['surface']
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20

    # Inicializar analizador
    analyzer = BackScheduleAnalyzer()

    # Contenedor principal
    main_container = ft.Column(spacing=20)

    # Estado para notificaciones
    export_status = ft.Text("", color=COLORS['text_secondary'])

    # Agregar contenedor principal UNA SOLA VEZ
    page.add(main_container)

    def export_data():
        """Exportar análisis"""
        export_status.value = "🔄 Exportando análisis corregido..."
        export_status.color = COLORS['warning']
        page.update()

        result = analyzer.export_effectiveness_analysis_to_excel()

        if result["success"]:
            export_status.value = f"✅ {result['message']}"
            export_status.color = COLORS['success']
        else:
            export_status.value = f"❌ {result['message']}"
            export_status.color = COLORS['error']

        page.update()

    def load_data():
        """Cargar datos"""
        analyzer.load_data_from_db()
        analyzer.load_fcst_alignment_data()
        analyzer.load_sales_order_alignment_data()
        analyzer.load_wo_materials_analysis()  # Nueva función WO-Materials
        analyzer.load_obsolete_expired_neteable_analysis()  # Nueva función OBS/EXP
        update_dashboard()

    def update_dashboard():
        """Actualizar dashboard con pestañas"""
        main_container.controls.clear()

        # Header
        header = ft.Container(
            content=ft.Row([
                ft.Column([
                    ft.Text("BackSchedule Analytics - INTERPRETACIÓN FINAL", 
                        size=32, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text(f"Análisis de {analyzer.metrics.get('total_records', 0):,} registros", 
                        size=16, color=COLORS['text_secondary']),
                    ft.Text("🔵 ShipDate - ReqDate | 🟢 ShipDate - ReqDate_Original", 
                        size=12, color=COLORS['accent'])
                ], expand=True),
                ft.Column([
                    ft.Row([
                        ft.ElevatedButton(
                            "🔄 Actualizar",
                            on_click=lambda _: load_data(),
                            bgcolor=COLORS['accent'],
                            color=COLORS['text_primary']
                        ),
                    ], spacing=10),
                    export_status
                ], horizontal_alignment=ft.CrossAxisAlignment.END)
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )

        # Crear contenido de la primera pestaña (Dashboard principal)
        def create_main_dashboard():
            gaussian_cards = create_gaussian_cards(analyzer)
            export_card = create_export_card(analyzer, lambda _: export_data())
            
            chart_base64 = analyzer.create_gaussian_chart()
            chart_container = ft.Container(
                content=ft.Column([
                    ft.Text("📊 ANÁLISIS FINAL CORREGIDO: Gestión de Abastecimiento", 
                        size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("🔵 BackSchedule (ShipDate - ReqDate) | 🟢 Sistema Original (ShipDate - ReqDate_Original)", 
                        size=12, color=COLORS['text_secondary']),
                    ft.Text("✅ Positivo = Material llega ANTES | ⭕ Cero = Material llega EXACTO | ❌ Negativo = Material llega DESPUÉS", 
                        size=11, color=COLORS['warning']),
                    ft.Image(
                        src_base64=chart_base64,
                        width=page.width - 80,
                        height=500,
                        fit=ft.ImageFit.CONTAIN
                    ) if chart_base64 else ft.Text("Error generando gráfico", color=COLORS['error'])
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )

            return ft.Column([
                gaussian_cards,
                export_card,
                chart_container
            ], spacing=20)

        def create_advanced_metrics(analyzer, page):
            # Cargar datos de FCST si no están cargados
            if not hasattr(analyzer, 'fcst_data'):
                analyzer.load_fcst_alignment_data()
            
            # Métricas de FCST
            metrics = analyzer.fcst_metrics
            
            # Tarjetas de métricas FCST
            fcst_cards = ft.Row([
                create_metric_card(
                    "Total Registros FCST", 
                    f"{metrics.get('total_records', 0):,}",
                    "OpenQty > 0 con WO",
                    COLORS['accent']
                ),
                create_metric_card(
                    "Alineación Perfecta", 
                    f"{metrics.get('alignment_pct', 0)}%",
                    f"{metrics.get('aligned_count', 0):,} registros",
                    COLORS['success']
                ),
                create_metric_card(
                    "Desalineados", 
                    f"{metrics.get('misaligned_count', 0):,}",
                    f"{100-metrics.get('alignment_pct', 0):.1f}%",
                    COLORS['error']
                ),
                create_metric_card(
                    "Diferencia Promedio", 
                    f"{metrics.get('mean_diff_days', 0)} días",
                    f"±{metrics.get('std_diff_days', 0)} std",
                    COLORS['warning']
                )
            ], wrap=True, spacing=15)

            # Gráfico de alineación
            chart_base64 = analyzer.create_fcst_alignment_chart()
            chart_container = ft.Container(
                content=ft.Column([
                    ft.Text("📊 ANÁLISIS DE ALINEACIÓN FCST vs WO", 
                        size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Comparación entre ReqDate (FCST) y DueDt (WO)", 
                        size=12, color=COLORS['text_secondary']),
                    ft.Image(
                        src_base64=chart_base64,
                        width=page.width - 80,
                        height=500,
                        fit=ft.ImageFit.CONTAIN
                    ) if chart_base64 else ft.Text("Error generando gráfico", color=COLORS['error'])
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )

            def export_fcst_analysis():
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    export_path = f"FCST_Alignment_Analysis_{timestamp}.xlsx"
                    
                    with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                        analyzer.fcst_data.to_excel(writer, sheet_name='FCST_Complete', index=False)
                        misaligned = analyzer.fcst_data[~analyzer.fcst_data['Is_Aligned']]
                        misaligned.to_excel(writer, sheet_name='FCST_Misaligned', index=False)
                        aligned = analyzer.fcst_data[analyzer.fcst_data['Is_Aligned']]
                        aligned.to_excel(writer, sheet_name='FCST_Aligned', index=False)
                    
                    # Abrir archivo
                    full_path = os.path.abspath(export_path)
                    import subprocess
                    import platform
                    if platform.system() == 'Windows':
                        subprocess.call(('start', full_path), shell=True)
                    
                    return f"✅ Exportado: {export_path}"
                except Exception as e:
                    return f"❌ Error: {str(e)}"
            
            export_fcst_card = ft.Card(
                content=ft.Container(
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12,
                    content=ft.Column([
                        ft.Row([
                            ft.Text("📤 Exportar Análisis FCST", 
                                    size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                            ft.ElevatedButton(
                                "📊 Exportar FCST",
                                on_click=lambda _: print(export_fcst_analysis()),
                                bgcolor=COLORS['accent'],
                                color=COLORS['text_primary']
                            )
                        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                        ft.Text("Se exportarán: Completos, Alineados, Desalineados", 
                                size=12, color=COLORS['text_secondary'])
                    ])
                ),
                elevation=8
            )

            return ft.Container(
                content=ft.Column([
                    ft.Text("📈 ANÁLISIS DE ALINEACIÓN FCST", 
                        size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Validación de fechas entre FCST (ReqDate) y WO (DueDt)", 
                        size=14, color=COLORS['text_secondary']),
                    fcst_cards,
                    export_fcst_card,
                    chart_container
                ], spacing=20),
                padding=ft.padding.all(20)
            )

        def create_so_wo_analysis(analyzer, page):
            # Cargar datos de SO-WO si no están cargados
            if not hasattr(analyzer, 'so_data'):
                analyzer.load_sales_order_alignment_data()
            
            # Métricas de SO-WO
            metrics = analyzer.so_metrics
            
            # Tarjetas de métricas SO-WO
            so_cards = ft.Row([
                create_metric_card(
                    "Total SO-WO", 
                    f"{metrics.get('total_records', 0):,}",
                    "Sales Orders con WO",
                    COLORS['accent']
                ),
                create_metric_card(
                    "Alineación Perfecta", 
                    f"{metrics.get('alignment_pct', 0)}%",
                    f"{metrics.get('aligned_count', 0):,} registros",
                    COLORS['success']
                ),
                create_metric_card(
                    "Keys Coincidentes", 
                    f"{metrics.get('matched_keys', 0):,}",
                    f"de {metrics.get('so_keys_unique', 0):,} SO únicas",
                    COLORS['warning']
                ),
                create_metric_card(
                    "Diferencia Promedio", 
                    f"{metrics.get('mean_diff_days', 0)} días",
                    f"DueDt vs Prd_Dt",
                    COLORS['error']
                )
            ], wrap=True, spacing=15)

            # Gráfico de alineación SO-WO
            chart_base64 = analyzer.create_so_alignment_chart()
            chart_container = ft.Container(
                content=ft.Column([
                    ft.Text("📊 ANÁLISIS DE ALINEACIÓN SALES ORDER vs WORK ORDER", 
                        size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Comparación entre Prd_Dt (SO) y DueDt (WO) usando llaves SO_No+Ln vs SO_FCST+Sub", 
                        size=12, color=COLORS['text_secondary']),
                    ft.Image(
                        src_base64=chart_base64,
                        width=page.width - 80,
                        height=500,
                        fit=ft.ImageFit.CONTAIN
                    ) if chart_base64 else ft.Text("Error generando gráfico", color=COLORS['error'])
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )

            def export_so_analysis():
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    export_path = f"SO_WO_Alignment_Analysis_{timestamp}.xlsx"
                    
                    with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                        # Exportar datos completos
                        analyzer.so_data.to_excel(writer, sheet_name='SO_WO_Complete', index=False)
                        
                        # Exportar por categorías
                        misaligned = analyzer.so_data[~analyzer.so_data['Is_Aligned']]
                        misaligned.to_excel(writer, sheet_name='SO_WO_Misaligned', index=False)
                        
                        aligned = analyzer.so_data[analyzer.so_data['Is_Aligned']]
                        aligned.to_excel(writer, sheet_name='SO_WO_Aligned', index=False)
                        
                        # WO tempranas y tardías
                        early_wo = analyzer.so_data[analyzer.so_data['Date_Diff_Days'] < 0]
                        early_wo.to_excel(writer, sheet_name='WO_Early', index=False)
                        
                        late_wo = analyzer.so_data[analyzer.so_data['Date_Diff_Days'] > 0]
                        late_wo.to_excel(writer, sheet_name='WO_Late', index=False)
                    
                    # Abrir archivo
                    full_path = os.path.abspath(export_path)
                    import subprocess
                    import platform
                    if platform.system() == 'Windows':
                        subprocess.call(('start', full_path), shell=True)
                    
                    return f"✅ Exportado: {export_path}"
                except Exception as e:
                    return f"❌ Error: {str(e)}"
            
            export_so_card = ft.Card(
                content=ft.Container(
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12,
                    content=ft.Column([
                        ft.Row([
                            ft.Text("📤 Exportar Análisis SO-WO", 
                                    size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                            ft.ElevatedButton(
                                "📊 Exportar SO-WO",
                                on_click=lambda _: print(export_so_analysis()),
                                bgcolor=COLORS['warning'],
                                color=COLORS['text_primary']
                            )
                        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                        ft.Text("Se exportarán: Completos, Alineados, Desalineados, WO Tempranas, WO Tardías", 
                                size=12, color=COLORS['text_secondary'])
                    ])
                ),
                elevation=8
            )

            # Tarjeta de resumen de llaves
            keys_summary_card = ft.Card(
                content=ft.Container(
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['secondary'],
                    border_radius=12,
                    content=ft.Column([
                        ft.Text("🔑 Resumen de Llaves de Cruce", 
                                size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                        ft.Divider(color=COLORS['primary']),
                        ft.Row([
                            ft.Column([
                                ft.Text("📋 Sales Orders", size=14, color=COLORS['accent']),
                                ft.Text("Llave: SO_No + Ln", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"Keys únicas: {metrics.get('so_keys_unique', 0):,}", 
                                        size=16, color=COLORS['text_primary'])
                            ], expand=1),
                            ft.Column([
                                ft.Text("🔧 Work Orders", size=14, color=COLORS['success']),
                                ft.Text("Llave: SO_FCST + Sub", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"Keys únicas: {metrics.get('wo_keys_unique', 0):,}", 
                                        size=16, color=COLORS['text_primary'])
                            ], expand=1),
                            ft.Column([
                                ft.Text("🎯 Coincidencias", size=14, color=COLORS['warning']),
                                ft.Text("Keys que coinciden", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"{metrics.get('matched_keys', 0):,}", 
                                        size=16, color=COLORS['text_primary'])
                            ], expand=1)
                        ])
                    ])
                ),
                elevation=8
            )

            return ft.Container(
                content=ft.Column([
                    ft.Text("🛒 ANÁLISIS DE ALINEACIÓN SALES ORDER vs WORK ORDER", 
                        size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Validación de fechas entre SO (Prd_Dt) y WO (DueDt) usando llaves compuestas", 
                        size=14, color=COLORS['text_secondary']),
                    so_cards,
                    keys_summary_card,
                    export_so_card,
                    chart_container
                ], spacing=20),
                padding=ft.padding.all(20)
            )

        def create_wo_materials_analysis(analyzer, page):
            # Cargar datos de WO-Materials si no están cargados
            if not hasattr(analyzer, 'wo_materials_data'):
                analyzer.load_wo_materials_analysis()
            
            # Métricas de WO-Materials
            metrics = analyzer.wo_materials_metrics
            
            # Tarjetas principales de métricas
            main_cards = ft.Row([
                create_metric_card(
                    "Total WO Analizadas", 
                    f"{metrics.get('total_wo', 0):,}",
                    "GF, GFA, KZ con OpnQ > 0",
                    COLORS['accent']
                ),
                create_metric_card(
                    "WO Con Materiales", 
                    f"{metrics.get('materials_pct', 0)}%",
                    f"{metrics.get('wo_with_materials', 0):,} Work Orders",
                    COLORS['success']
                ),
                create_metric_card(
                    "WO Sin Materiales", 
                    f"{metrics.get('without_materials_pct', 0)}%",
                    f"{metrics.get('wo_without_materials', 0):,} Work Orders",
                    COLORS['error']
                ),

            ], wrap=True, spacing=15)

            # Tarjetas por tipo Srt
            srt_cards = ft.Row([
                ft.Card(
                    content=ft.Container(
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['primary'],
                        border_radius=10,
                        content=ft.Column([
                            ft.Row([
                                ft.Text("🔧", size=20),
                                ft.Text("Tipo GF", size=16, color=COLORS['accent'], weight=ft.FontWeight.BOLD)
                            ]),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Row([
                                ft.Column([
                                    ft.Text("Total", size=12, color=COLORS['text_secondary']),
                                    ft.Text(f"{metrics.get('srt_gf_total', 0):,}", size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)
                                ], expand=1),
                                ft.Column([
                                    ft.Text("Sin Materiales", size=12, color=COLORS['text_secondary']),
                                    ft.Text(f"{metrics.get('srt_gf_without', 0):,}", size=18, color=COLORS['error'], weight=ft.FontWeight.BOLD)
                                ], expand=1)
                            ]),
                            ft.Text(f"{100*metrics.get('srt_gf_without', 0)/max(metrics.get('srt_gf_total', 1), 1):.1f}% sin materiales", 
                                    size=11, color=COLORS['text_secondary'])
                        ], spacing=8)
                    ),
                    elevation=6
                ),
                ft.Card(
                    content=ft.Container(
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['primary'],
                        border_radius=10,
                        content=ft.Column([
                            ft.Row([
                                ft.Text("⚙️", size=20),
                                ft.Text("Tipo GFA", size=16, color=COLORS['success'], weight=ft.FontWeight.BOLD)
                            ]),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Row([
                                ft.Column([
                                    ft.Text("Total", size=12, color=COLORS['text_secondary']),
                                    ft.Text(f"{metrics.get('srt_gfa_total', 0):,}", size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)
                                ], expand=1),
                                ft.Column([
                                    ft.Text("Sin Materiales", size=12, color=COLORS['text_secondary']),
                                    ft.Text(f"{metrics.get('srt_gfa_without', 0):,}", size=18, color=COLORS['error'], weight=ft.FontWeight.BOLD)
                                ], expand=1)
                            ]),
                            ft.Text(f"{100*metrics.get('srt_gfa_without', 0)/max(metrics.get('srt_gfa_total', 1), 1):.1f}% sin materiales", 
                                    size=11, color=COLORS['text_secondary'])
                        ], spacing=8)
                    ),
                    elevation=6
                ),
                ft.Card(
                    content=ft.Container(
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['primary'],
                        border_radius=10,
                        content=ft.Column([
                            ft.Row([
                                ft.Text("🔩", size=20),
                                ft.Text("Tipo KZ", size=16, color=COLORS['warning'], weight=ft.FontWeight.BOLD)
                            ]),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Row([
                                ft.Column([
                                    ft.Text("Total", size=12, color=COLORS['text_secondary']),
                                    ft.Text(f"{metrics.get('srt_kz_total', 0):,}", size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)
                                ], expand=1),
                                ft.Column([
                                    ft.Text("Sin Materiales", size=12, color=COLORS['text_secondary']),
                                    ft.Text(f"{metrics.get('srt_kz_without', 0):,}", size=18, color=COLORS['error'], weight=ft.FontWeight.BOLD)
                                ], expand=1)
                            ]),
                            ft.Text(f"{100*metrics.get('srt_kz_without', 0)/max(metrics.get('srt_kz_total', 1), 1):.1f}% sin materiales", 
                                    size=11, color=COLORS['text_secondary'])
                        ], spacing=8)
                    ),
                    elevation=6
                )
            ], spacing=15)

            # Gráficos de análisis
            chart_base64 = analyzer.create_wo_materials_chart()
            chart_container = ft.Container(
                content=ft.Column([
                    ft.Text("📊 ANÁLISIS COMPLETO: WO SIN MATERIALES", 
                        size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Work Orders (GF/GFA/KZ) vs Materiales en PR561", 
                        size=12, color=COLORS['text_secondary']),
                    ft.Image(
                        src_base64=chart_base64,
                        width=page.width - 80,
                        height=800,  # Mayor altura para 4 gráficos
                        fit=ft.ImageFit.CONTAIN
                    ) if chart_base64 else ft.Text("Error generando gráfico", color=COLORS['error'])
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )

            # Función de exportación
            def export_wo_materials_analysis():
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    export_path = f"WO_Materials_Analysis_{timestamp}.xlsx"
                    
                    with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                        # Exportar datos principales
                        analyzer.wo_materials_data.to_excel(writer, sheet_name='WO_Complete', index=False)
                        
                        # Exportar por categorías
                        wo_with_materials = analyzer.wo_materials_data[analyzer.wo_materials_data['Has_Materials']]
                        wo_with_materials.to_excel(writer, sheet_name='WO_With_Materials', index=False)
                        
                        wo_without_materials = analyzer.wo_materials_data[~analyzer.wo_materials_data['Has_Materials']]
                        wo_without_materials.to_excel(writer, sheet_name='WO_Without_Materials', index=False)
                        
                        # Exportar por tipo Srt sin materiales
                        for srt_type in ['GF', 'GFA', 'KZ']:
                            srt_without = wo_without_materials[wo_without_materials['Srt'] == srt_type]
                            if len(srt_without) > 0:
                                srt_without.to_excel(writer, sheet_name=f'{srt_type}_Without_Materials', index=False)
                        
                        # Exportar detalles de materiales (pr561)
                        if hasattr(analyzer, 'wo_materials_raw') and analyzer.wo_materials_raw is not None:
                            analyzer.wo_materials_raw.to_excel(writer, sheet_name='PR561_Details', index=False)
                        
                        # Crear hoja de resumen
                        summary_data = {
                            'Métrica': [
                                'Total WO Analizadas',
                                'WO Con Materiales',
                                'WO Sin Materiales',
                                '% Con Materiales',
                                '% Sin Materiales',
                                'Promedio Materiales por WO',
                                'GF Total',
                                'GF Sin Materiales',
                                'GFA Total', 
                                'GFA Sin Materiales',
                                'KZ Total',
                                'KZ Sin Materiales'
                            ],
                            'Valor': [
                                metrics.get('total_wo', 0),
                                metrics.get('wo_with_materials', 0),
                                metrics.get('wo_without_materials', 0),
                                f"{metrics.get('materials_pct', 0)}%",
                                f"{metrics.get('without_materials_pct', 0)}%",
                                metrics.get('avg_materials_per_wo', 0),
                                metrics.get('srt_gf_total', 0),
                                metrics.get('srt_gf_without', 0),
                                metrics.get('srt_gfa_total', 0),
                                metrics.get('srt_gfa_without', 0),
                                metrics.get('srt_kz_total', 0),
                                metrics.get('srt_kz_without', 0)
                            ]
                        }
                        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary_Metrics', index=False)
                    
                    # Abrir archivo
                    full_path = os.path.abspath(export_path)
                    import subprocess
                    import platform
                    if platform.system() == 'Windows':
                        subprocess.call(('start', full_path), shell=True)
                    
                    return f"✅ Exportado: {export_path}"
                except Exception as e:
                    return f"❌ Error: {str(e)}"

            # Tarjeta de exportación
            export_materials_card = ft.Card(
                content=ft.Container(
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12,
                    content=ft.Column([
                        ft.Row([
                            ft.Text("📤 Exportar Análisis WO-Materials", 
                                    size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                            ft.ElevatedButton(
                                "📊 Exportar WO-Materials",
                                on_click=lambda _: print(export_wo_materials_analysis()),
                                bgcolor=COLORS['error'],
                                color=COLORS['text_primary']
                            )
                        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                        ft.Text("Se exportarán: Completas, Con/Sin Materiales, Por Tipo Srt, Detalles PR561, Resumen", 
                                size=12, color=COLORS['text_secondary'])
                    ])
                ),
                elevation=8
            )

            # Tarjeta de información del análisis
            info_card = ft.Card(
                content=ft.Container(
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['secondary'],
                    border_radius=12,
                    content=ft.Column([
                        ft.Text("ℹ️ Información del Análisis", 
                                size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                        ft.Divider(color=COLORS['primary']),
                        ft.Column([
                            ft.Text("🔍 Filtros Aplicados:", size=14, color=COLORS['accent'], weight=ft.FontWeight.BOLD),
                            ft.Text("• Work Orders con Srt IN ('GF', 'GFA', 'KZ')", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Work Orders con OpnQ > 0 (cantidad abierta)", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Cruce con tabla PR561 por WONo", size=12, color=COLORS['text_secondary']),
                            ft.Text("", size=4),
                            ft.Text("📊 Definiciones:", size=14, color=COLORS['success'], weight=ft.FontWeight.BOLD),
                            ft.Text("• Con Materiales: WO que tiene al menos 1 registro en PR561", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Sin Materiales: WO que NO tiene registros en PR561", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Esto indica WO sin bill of materials cargado", size=12, color=COLORS['text_secondary']),
                        ], spacing=3)
                    ])
                ),
                elevation=8
            )

            return ft.Container(
                content=ft.Column([
                    ft.Text("🔧 ANÁLISIS: WORK ORDERS SIN MATERIALES", 
                        size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Identificación de WO (GF/GFA/KZ) sin materiales cargados en PR561", 
                        size=14, color=COLORS['text_secondary']),
                    main_cards,
                    srt_cards,
                    # Tarjetas por Status
                    ft.Text("📊 ANÁLISIS POR STATUS DE WORK ORDER", 
                            size=18, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    create_status_cards(metrics),
                    ft.Divider(color=COLORS['secondary']),
                    info_card,
                    export_materials_card,
                    chart_container
                ], spacing=20),
                padding=ft.padding.all(20)
            )
        
        def create_obsolete_expired_analysis(analyzer, page):
            """Crear análisis de materiales obsoletos/expirados en localidades neteables"""
            # Cargar datos de OBS/EXP si no están cargados
            if not hasattr(analyzer, 'obsolete_data'):
                analyzer.load_obsolete_expired_neteable_analysis()
            
            # Métricas de OBS/EXP
            metrics = analyzer.obsolete_metrics
            
            # Tarjetas principales de métricas
            main_cards = ft.Row([
                create_metric_card(
                    "Total en Loc. Neteables", 
                    f"{metrics.get('total_materials_neteable', 0):,}",
                    "Materiales con QtyOH > 0",
                    COLORS['accent']
                ),
                create_metric_card(
                    "Materiales Expirados", 
                    f"{metrics.get('expired_percentage', 0)}%",
                    f"{metrics.get('total_expired', 0):,} materiales",
                    COLORS['error']
                ),
                create_metric_card(
                    "Items Únicos Expirados", 
                    f"{metrics.get('unique_expired_items', 0):,}",
                    f"En {metrics.get('locations_with_expired', 0):,} localidades",
                    COLORS['success']
                )
            ], wrap=True, spacing=15)

            # Resto de la función... (es muy larga, ¿quieres que te dé solo el inicio para que confirmes la ubicación correcta?)
            # Tarjetas por severidad de expiración
            severity_cards = ft.Row([
                ft.Card(
                    content=ft.Container(
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['primary'],
                        border_radius=10,
                        content=ft.Column([
                            ft.Row([
                                ft.Text("🚨", size=20),
                                ft.Text("Crítico", size=16, color='#8B0000', weight=ft.FontWeight.BOLD)
                            ]),
                            ft.Text("<= 0%", size=12, color=COLORS['text_secondary']),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Text(f"{metrics.get('critico_count', 0):,}", size=24, color='#8B0000', weight=ft.FontWeight.BOLD),
                            ft.Text("materiales", size=12, color=COLORS['text_secondary'])
                        ], spacing=6)
                    ),
                    elevation=6
                ),
                ft.Card(
                    content=ft.Container(
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['primary'],
                        border_radius=10,
                        content=ft.Column([
                            ft.Row([
                                ft.Text("⚠️", size=20),
                                ft.Text("Muy Expirado", size=16, color='#FF4500', weight=ft.FontWeight.BOLD)
                            ]),
                            ft.Text("-180 a -365 días", size=12, color=COLORS['text_secondary']),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Text(f"{metrics.get('muy_expirado_count', 0):,}", size=24, color='#FF4500', weight=ft.FontWeight.BOLD),
                            ft.Text("materiales", size=12, color=COLORS['text_secondary'])
                        ], spacing=6)
                    ),
                    elevation=6
                ),
                ft.Card(
                    content=ft.Container(
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['primary'],
                        border_radius=10,
                        content=ft.Column([
                            ft.Row([
                                ft.Text("🔶", size=20),
                                ft.Text("Expirado", size=16, color='#FF6347', weight=ft.FontWeight.BOLD)
                            ]),
                            ft.Text("-30 a -180 días", size=12, color=COLORS['text_secondary']),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Text(f"{metrics.get('expirado_count', 0):,}", size=24, color='#FF6347', weight=ft.FontWeight.BOLD),
                            ft.Text("materiales", size=12, color=COLORS['text_secondary'])
                        ], spacing=6)
                    ),
                    elevation=6
                ),
                ft.Card(
                    content=ft.Container(
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['primary'],
                        border_radius=10,
                        content=ft.Column([
                            ft.Row([
                                ft.Text("🟡", size=20),
                                ft.Text("Recién Expirado", size=16, color='#FFB366', weight=ft.FontWeight.BOLD)
                            ]),
                            ft.Text("0 a -30 días", size=12, color=COLORS['text_secondary']),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Text(f"{metrics.get('recien_expirado_count', 0):,}", size=24, color='#FFB366', weight=ft.FontWeight.BOLD),
                            ft.Text("materiales", size=12, color=COLORS['text_secondary'])
                        ], spacing=6)
                    ),
                    elevation=6
                )
            ], spacing=15)

            # Gráficos de análisis
            chart_base64 = analyzer.create_obsolete_expired_chart()
            chart_container = ft.Container(
                content=ft.Column([
                    ft.Text("📊 ANÁLISIS COMPLETO: MATERIALES OBSOLETOS/EXPIRADOS EN LOCALIDADES NETEABLES", 
                        size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Shelf_Life <= 0 en localidades con TagNetable = 'YES'", 
                        size=12, color=COLORS['text_secondary']),
                    ft.Image(
                        src_base64=chart_base64,
                        width=page.width - 80,
                        height=800,
                        fit=ft.ImageFit.CONTAIN
                    ) if chart_base64 else ft.Text("Error generando gráfico", color=COLORS['error'])
                ]),
                padding=ft.padding.all(20),
                bgcolor=COLORS['primary'],
                border_radius=12
            )

            # Top 10 Items más críticos
            top_items_card = ft.Card(
                content=ft.Container(
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['secondary'],
                    border_radius=12,
                    content=ft.Column([
                        ft.Text("🏆 TOP 10 ITEMS MÁS CRÍTICOS (Por Cantidad)", 
                                size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                        ft.Divider(color=COLORS['primary']),
                        create_top_items_table(analyzer)
                    ])
                ),
                elevation=8
            )

            # Función de exportación
            def export_obsolete_analysis():
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    export_path = f"Obsolete_Expired_Neteable_Analysis_{timestamp}.xlsx"
                    
                    # VERIFICAR datos antes de exportar
                    print(f"🔍 VERIFICACIÓN antes de exportar:")
                    print(f"   Total materiales: {len(analyzer.obsolete_data)}")
                    print(f"   Materiales expirados: {len(analyzer.expired_data)}")
                    print(f"   Rango Shelf_Life en expirados: {analyzer.expired_data['Shelf_Life'].min():.2f}% a {analyzer.expired_data['Shelf_Life'].max():.2f}%")
                    print(f"   Warehouses en expirados: {sorted(analyzer.expired_data['Wh'].unique())}")
                    print(f"   Zonas en expirados: {sorted(analyzer.expired_data['Zone'].unique())}")
                    
                    with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                        # Exportar datos completos (NORMALIZADOS)
                        all_normalized = analyzer.obsolete_data.copy()
                        all_normalized['Wh'] = all_normalized['Wh'].astype(str).str.upper().str.strip()
                        all_normalized['Zone'] = all_normalized['Zone'].astype(str).str.upper().str.strip()
                        all_normalized.to_excel(writer, sheet_name='All_Materials_Neteable', index=False)
                        
                        # Exportar solo expirados (NORMALIZADOS)
                        expired_normalized = analyzer.expired_data.copy()
                        expired_normalized['Wh'] = expired_normalized['Wh'].astype(str).str.upper().str.strip()
                        expired_normalized['Zone'] = expired_normalized['Zone'].astype(str).str.upper().str.strip()
                        expired_normalized.to_excel(writer, sheet_name='Expired_Materials', index=False)
                        
                        # Exportar por categorías de severidad
                        for category in ['Crítico (<= 5%)', 'Muy Bajo (5-10%)', 'Bajo (10-25%)', 'Medio (25-50%)']:
                            cat_data = expired_normalized[expired_normalized['Expiry_Category'] == category]
                            if len(cat_data) > 0:
                                sheet_name = category.replace(' ', '_').replace('(', '').replace(')', '').replace('<', 'lt').replace('>', 'gt').replace('=', '').replace('%', 'pct').replace('-', '_')[:31]
                                cat_data.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Top items expirados (NORMALIZADO)
                        if hasattr(analyzer, 'top_expired_items') and analyzer.top_expired_items is not None:
                            analyzer.top_expired_items.to_excel(writer, sheet_name='Top_Expired_Items', index=False)
                        
                        # Resumen por warehouse (CORREGIDO)
                        wh_summary = expired_normalized.groupby('Wh').agg({
                            'ItemNo': 'count',
                            'QtyOH': 'sum',
                            'Shelf_Life': 'mean'
                        }).rename(columns={
                            'ItemNo': 'Items_Count',
                            'QtyOH': 'Total_Qty_Expired',
                            'Shelf_Life': 'Avg_Shelf_Life'
                        }).reset_index().sort_values('Total_Qty_Expired', ascending=False)
                        wh_summary.to_excel(writer, sheet_name='Summary_by_Warehouse', index=False)
                        
                        # Resumen por zona (CORREGIDO)
                        zone_summary = expired_normalized.groupby('Zone').agg({
                            'ItemNo': 'count',
                            'QtyOH': 'sum',
                            'Shelf_Life': 'mean'
                        }).rename(columns={
                            'ItemNo': 'Items_Count',
                            'QtyOH': 'Total_Qty_Expired',
                            'Shelf_Life': 'Avg_Shelf_Life'
                        }).reset_index().sort_values('Total_Qty_Expired', ascending=False)
                        zone_summary.to_excel(writer, sheet_name='Summary_by_Zone', index=False)
                        
                        # Crear hoja de resumen de métricas (ACTUALIZADA)
                        summary_data = {
                            'Métrica': [
                                'Total Materiales en Loc. Neteables',
                                'Materiales con Shelf_Life <= 50%',
                                '% Materiales con Shelf_Life <= 50%',
                                'Cantidad Total de Stock (Shelf_Life <= 50%)',
                                'Items Únicos con Shelf_Life <= 50%',
                                'Localidades con Materiales <= 50%',
                                'Warehouses Únicos',
                                'Zonas Únicas',
                                'Críticos (<= 5%)',
                                'Muy Bajo (5-10%)',
                                'Bajo (10-25%)',
                                'Medio (25-50%)',
                                'Top Warehouse',
                                'Top Zone',
                                'Threshold Usado'
                            ],
                            'Valor': [
                                metrics.get('total_materials_neteable', 0),
                                metrics.get('total_expired', 0),
                                f"{metrics.get('expired_percentage', 0)}%",
                                metrics.get('total_expired_qty', 0),
                                metrics.get('unique_expired_items', 0),
                                metrics.get('locations_with_expired', 0),
                                metrics.get('unique_warehouses', 0),
                                metrics.get('unique_zones', 0),
                                metrics.get('critico_count', 0),
                                metrics.get('muy_expirado_count', 0),
                                metrics.get('expirado_count', 0),
                                metrics.get('recien_expirado_count', 0),
                                f"{metrics.get('top_warehouse', 'N/A')} ({metrics.get('top_warehouse_count', 0)})",
                                f"{metrics.get('top_zone', 'N/A')} ({metrics.get('top_zone_count', 0)})",
                                "Shelf_Life <= 50%"
                            ]
                        }
                        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary_Metrics', index=False)
                    
                    # Abrir archivo
                    full_path = os.path.abspath(export_path)
                    import subprocess
                    import platform
                    if platform.system() == 'Windows':
                        subprocess.call(('start', full_path), shell=True)
                    
                    return f"✅ Exportado: {export_path}"
                except Exception as e:
                    return f"❌ Error: {str(e)}"










            # Tarjeta de exportación
            export_obsolete_card = ft.Card(
                content=ft.Container(
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12,
                    content=ft.Column([
                        ft.Row([
                            ft.Text("📤 Exportar Análisis OBS/EXP LOC NET", 
                                    size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                            ft.ElevatedButton(
                                "📊 Exportar OBS/EXP",
                                on_click=lambda _: print(export_obsolete_analysis()),
                                bgcolor=COLORS['error'],
                                color=COLORS['text_primary']
                            )
                        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                        ft.Text("Se exportarán: Todos, Expirados, Por Severidad, Top Items, Resúmenes por WH/Zona", 
                                size=12, color=COLORS['text_secondary'])
                    ])
                ),
                elevation=8
            )

            # Tarjeta de información del análisis
            info_card = ft.Card(
                content=ft.Container(
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['secondary'],
                    border_radius=12,
                    content=ft.Column([
                        ft.Text("ℹ️ Información del Análisis OBS/EXP LOC NET", 
                                size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                        ft.Divider(color=COLORS['primary']),
                        ft.Column([
                            ft.Text("🔍 Filtros Aplicados:", size=14, color=COLORS['accent'], weight=ft.FontWeight.BOLD),
                            ft.Text("• Localidades con TagNetable = 'YES'", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Materiales con QtyOH > 0 (cantidad en stock)", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Cruce por Bin (in521) = BinID (whs_location_in36851)", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Materiales con Shelf_Life <= 0 (expirados/obsoletos)", size=12, color=COLORS['text_secondary']),
                            ft.Text("", size=4),
                            ft.Text("📊 Definiciones:", size=14, color=COLORS['success'], weight=ft.FontWeight.BOLD),
                            ft.Text("• Localidad Neteable: TagNetable = 'YES' (puede ser reasignada)", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Material Expirado: Shelf_Life <= 0 días", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Crítico: < -365 días (más de 1 año expirado)", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Muy Expirado: -180 a -365 días", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Expirado: -30 a -180 días", size=12, color=COLORS['text_secondary']),
                            ft.Text("• Recién Expirado: 0 a -30 días", size=12, color=COLORS['text_secondary']),
                        ], spacing=3)
                    ])
                ),
                elevation=8
            )

            return ft.Container(
                content=ft.Column([
                    ft.Text("🏭 ANÁLISIS: MATERIALES OBSOLETOS/EXPIRADOS EN LOCALIDADES NETEABLES", 
                        size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Identificación de materiales con Shelf_Life <= 50% en localidades TagNetable = 'YES'", 
                        size=14, color=COLORS['text_secondary']),
                    main_cards,
                    severity_cards,
                    ft.Divider(color=COLORS['secondary']),
                    info_card,
                    top_items_card,
                    export_obsolete_card,
                    chart_container
                ], spacing=20),
                padding=ft.padding.all(20)
            )


        def create_status_cards(metrics):
            """Crear tarjetas dinámicas por WO_Status"""
            status_cards = []
            
            # Obtener todos los status únicos de las métricas
            status_list = []
            for key in metrics.keys():
                if key.startswith('status_') and key.endswith('_total'):
                    status = key.replace('status_', '').replace('_total', '')
                    status_list.append(status)
            
            # Crear tarjeta para cada status
            for status in sorted(status_list):
                total = metrics.get(f'status_{status}_total', 0)
                without = metrics.get(f'status_{status}_without', 0)
                without_pct = metrics.get(f'status_{status}_without_pct', 0)
                
                # Definir colores por status (puedes personalizar estos colores)
                color_map = {
                    'F': COLORS['error'],      # Rojo para Firme
                    'O': COLORS['warning'],    # Amarillo para Open
                    'R': COLORS['accent'],     # Azul para Released
                    'C': COLORS['success'],    # Verde para Completed
                }
                
                status_color = color_map.get(status, COLORS['text_secondary'])
                
                card = ft.Card(
                    content=ft.Container(
                        padding=ft.padding.all(15),
                        bgcolor=COLORS['primary'],
                        border_radius=10,
                        content=ft.Column([
                            ft.Row([
                                ft.Text("📋", size=18),
                                ft.Text(f"Status {status}", size=14, color=status_color, weight=ft.FontWeight.BOLD)
                            ]),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Row([
                                ft.Column([
                                    ft.Text("Total", size=11, color=COLORS['text_secondary']),
                                    ft.Text(f"{total:,}", size=16, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)
                                ], expand=1),
                                ft.Column([
                                    ft.Text("Sin Mat.", size=11, color=COLORS['text_secondary']),
                                    ft.Text(f"{without:,}", size=16, color=COLORS['error'], weight=ft.FontWeight.BOLD)
                                ], expand=1)
                            ]),
                            ft.Text(f"{without_pct}% sin materiales", 
                                    size=10, color=COLORS['text_secondary'], text_align=ft.TextAlign.CENTER)
                        ], spacing=6)
                    ),
                    elevation=4
                )
                status_cards.append(card)
            
            # Organizar en filas de máximo 4 tarjetas
            rows = []
            for i in range(0, len(status_cards), 4):
                row_cards = status_cards[i:i+4]
                rows.append(ft.Row(row_cards, spacing=10, wrap=True))
            
            return ft.Column(rows, spacing=15)


        # Crear pestañas
        tabs = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=[
                ft.Tab(
                    text="📊 Dashboard Principal",
                    content=create_main_dashboard()
                ),
                ft.Tab(
                    text="📈 Métricas FCST", 
                    content=create_advanced_metrics(analyzer, page)
                ),
                ft.Tab(
                    text="🛒 Análisis SO-WO",
                    content=create_so_wo_analysis(analyzer, page)
                ),
                ft.Tab(
                    text="🔧 WO sin Materiales",
                    content=create_wo_materials_analysis(analyzer, page)
                ),
                ft.Tab(
                    text="🏭 OBS/EXP LOC NET",
                    content=create_obsolete_expired_analysis(analyzer, page)
                )
            ],
            expand=1
        )

        # Agregar header y tabs
        main_container.controls.append(header)
        main_container.controls.append(tabs)
        page.update()

    # Cargar datos iniciales
    load_data()

if __name__ == "__main__":
    ft.app(target=main)