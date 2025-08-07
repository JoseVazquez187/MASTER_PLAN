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

    def load_fcst_alignment_data(self, db_path=r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"):
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

    # Agrega estos métodos a la clase BackScheduleAnalyzer

    def load_so_wo_alignment_data(self, db_path=r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"):
        """Cargar y validar alineación de Sales Order con WO"""
        try:
            print(f"Conectando a R4Database para SO-WO: {db_path}")
            conn = sqlite3.connect(db_path)
            
            # Query para Sales Order - SIN FILTRAR POR WO_No, solo por Opn_Q
            so_query = """
            SELECT Entity, Proj, SO_No, Ln, Ord_Cd, Spr_CD, Order_Dt, Cancel_Dt, 
                CustReqDt, Req_Dt, Pln_Ship, Prd_Dt, Inv_Dt, Recover_Dt, 
                TimeToConfirm_hr, Std_LT, Cust, Cust_Name, Type_Code, Cust_PO, 
                Buyer_Name, AC, Item_Number, Description, Item_Rev, UM, ML, 
                Orig_Q, Opn_Q, OH_Netable, OH_Non_Netable, TR, Issue_Q, 
                Plnr_Intl, Plnr_Code, PlanType, Family, Misc_Code, WO_No, 
                WO_Due_dt, Unit_Price, OpenValue, Curr, T_Desc, So_Ln_Memo, 
                WO_Memo, ST, AC_Priority, T_Card, Std_Cost, Cust_NCR, 
                Reason_Cd, User_ID, Invoice_No, Shipped_Date, Ship_no, 
                Ship_via, Tracking_No, Vendor_Code, Vendor_Name, DG, 
                Planner_Notes, Responsibility, Last_Notes, Manuf_Charge, 
                Memo_1, Memo_2
            FROM sales_order_table
            WHERE Opn_Q > 0
            """
            
            # Query para WOInquiry - SOLO las que tienen SO_FCST
            wo_query = """
            SELECT id, Entity, ProjectNo, WONo, SO_FCST, Sub, ParentWO, 
                ItemNumber, Rev, Description, MiscCode, AC, LeadTime, DueDt,
                A_St_Dt, CompDt, CreateDt, WoType, Srt, Plnr, PlanType,
                ItmTp, UM, IssQ, CompQ, OpnQ, St, QAAprvl, Stk, Iss, Prt,
                UserId, StdCost, PrtNo, PrtUser, PrtDate, OpnHrs, SumLaborHr,
                StaticPlanNo, SPRev, StaticPlanDesc, WOLastNotes
            FROM WOInquiry
            WHERE SO_FCST IS NOT NULL AND SO_FCST != "" AND Sub IS NOT NULL
            """
            # 2. SEGUNDO: Cargar los datos
            so_df = pd.read_sql_query(so_query, conn)
            wo_df = pd.read_sql_query(wo_query, conn)
            conn.close()
        
            print(f"✅ Sales Order cargado: {len(so_df)} registros (con Opn_Q > 0)")
            print(f"✅ WOInquiry cargado: {len(wo_df)} registros (con SO_FCST)")
            
            # 3. TERCERO: Verificar que tenemos datos
            if len(so_df) == 0:
                print("❌ No hay registros de Sales Order con Opn_Q > 0")
                self._create_sample_so_wo_data()
                return False
                
            if len(wo_df) == 0:
                print("❌ No hay registros de WO con SO_FCST")
                self._create_sample_so_wo_data()
                return False
            

            # Crear llaves compuestas ANTES de filtrar
            print(f"🔑 Creando llaves compuestas...")
            so_df['SO_Key'] = so_df['SO_No'].astype(str) + '-' + so_df['Ln'].astype(str)
            wo_df['WO_Key'] = wo_df['SO_FCST'].astype(str) + '-' + wo_df['Sub'].astype(str)
            
            print(f"🔑 Llaves SO creadas: {so_df['SO_Key'].nunique()} únicas de {len(so_df)} registros")
            print(f"🔑 Llaves WO creadas: {wo_df['WO_Key'].nunique()} únicas de {len(wo_df)} registros")
            
            # AHORA SÍ hacer el cruce para ver coincidencias
            print(f"🔗 Buscando coincidencias entre SO y WO...")
            common_keys = set(so_df['SO_Key']).intersection(set(wo_df['WO_Key']))
            print(f"🎯 Llaves que coinciden: {len(common_keys)}")
            
            if len(common_keys) > 0:
                print(f"📋 Ejemplos de llaves coincidentes: {list(common_keys)[:5]}")
            else:
                print(f"❌ NO HAY COINCIDENCIAS - Verificando ejemplos:")
                print(f"📋 Primeras 5 llaves SO: {so_df['SO_Key'].head().tolist()}")
                print(f"📋 Primeras 5 llaves WO: {wo_df['WO_Key'].head().tolist()}")
            
            # Procesar fechas
            so_df["Prd_Dt"] = pd.to_datetime(so_df["Prd_Dt"], errors='coerce')
            wo_df["DueDt"] = pd.to_datetime(wo_df["DueDt"], errors='coerce')
            
            # Cruzar tablas: SO_Key = WO_Key (INNER JOIN para solo coincidencias)
            merged_df = so_df.merge(wo_df, left_on='SO_Key', right_on='WO_Key', how='inner', suffixes=('_so', '_wo'))
            
            print(f"✅ Registros cruzados: {len(merged_df)}")
            print(f"📊 Coincidencias encontradas: {len(merged_df)} de {len(so_df)} SO")
            
            # Validar alineación: Prd_Dt vs DueDt
            merged_df = merged_df.dropna(subset=['Prd_Dt', 'DueDt'])
            merged_df['Date_Diff_Days'] = (merged_df['DueDt'] - merged_df['Prd_Dt']).dt.days
            merged_df['Is_Aligned'] = merged_df['Date_Diff_Days'] == 0
            
            # Análisis de diferencias
            merged_df['Alignment_Status'] = merged_df.apply(lambda row: 
                'Aligned' if row['Date_Diff_Days'] == 0 
                else 'WO_Early' if row['Date_Diff_Days'] < 0 
                else 'WO_Late', axis=1)
            
            # Métricas de alineación
            total_records = len(merged_df)
            aligned_count = merged_df['Is_Aligned'].sum()
            wo_early_count = (merged_df['Date_Diff_Days'] < 0).sum()  # WO vence antes que SO
            wo_late_count = (merged_df['Date_Diff_Days'] > 0).sum()   # WO vence después que SO
            alignment_pct = (aligned_count / total_records * 100) if total_records > 0 else 0
            
            # Estadísticas de diferencias
            mean_diff = merged_df['Date_Diff_Days'].mean()
            std_diff = merged_df['Date_Diff_Days'].std()
            median_diff = merged_df['Date_Diff_Days'].median()
            
            # Análisis por entidad y AC
            entity_analysis = merged_df.groupby('Entity_so').agg({
                'Is_Aligned': ['count', 'sum'],
                'Date_Diff_Days': ['mean', 'std']
            }).round(2)
            
            ac_analysis = merged_df.groupby('AC_so').agg({
                'Is_Aligned': ['count', 'sum'], 
                'Date_Diff_Days': ['mean', 'std']
            }).round(2)
            
            # Guardar resultados
            self.so_wo_data = merged_df
            self.so_wo_metrics = {
                'total_records': total_records,
                'aligned_count': int(aligned_count),
                'wo_early_count': int(wo_early_count),
                'wo_late_count': int(wo_late_count),
                'alignment_pct': round(alignment_pct, 2),
                'mean_diff_days': round(mean_diff, 2),
                'std_diff_days': round(std_diff, 2),
                'median_diff_days': round(median_diff, 2),
                'so_total': len(so_df),
                'wo_total': len(wo_df),
                'match_rate': round(100 * total_records / len(so_df), 2) if len(so_df) > 0 else 0  # CORREGIR DIVISIÓN POR CERO
            }
            
            return True
            
        except Exception as e:
            print(f"❌ Error cargando SO-WO: {e}")
            # Crear datos de ejemplo para testing
            self._create_sample_so_wo_data()
            return False

    def _create_sample_so_wo_data(self):
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
            
            prd_dt = base_date + timedelta(days=i + np.random.randint(-5, 45))
            
            # Simular alineación: 60% alineado, 40% desalineado
            if np.random.random() < 0.6:
                due_dt = prd_dt  # Perfectamente alineado
            else:
                # Desalineado con variaciones más realistas para SO-WO
                diff_days = int(np.random.choice([-15, -7, -3, -1, 1, 3, 7, 15, 21], 
                                            p=[0.05, 0.10, 0.15, 0.20, 0.20, 0.15, 0.10, 0.04, 0.01]))  # CONVERTIR A INT
                due_dt = prd_dt + timedelta(days=diff_days)
            
            record = {
                # SO Data
                'Entity_so': f'Entity_{i%3}',
                'SO_No': so_no,
                'Ln': ln,
                'SO_Key': so_key,
                'Item_Number': f'SO-Item-{100+i:03d}',
                'Description_so': f'Sales Order Item {i}',
                'Prd_Dt': prd_dt,
                'Opn_Q': np.random.randint(1, 100),
                'Cust_Name': f'Customer_{i%10}',
                'AC_so': f'AC{i%4}',
                'WO_No': f'WO-{3000+i:04d}',
                # WO Data
                'WONo': f'WO-{3000+i:04d}',
                'WO_Key': so_key,
                'SO_FCST': so_no,
                'Sub': ln,
                'DueDt': due_dt,
                'Description_wo': f'WO Item {i}',
                'WoType': np.random.choice(['Make', 'Buy', 'Transfer']),
                'AC_wo': f'AC{i%4}',
                'St': np.random.choice(['Open', 'Released', 'Complete']),
                # Calculated
                'Date_Diff_Days': (due_dt - prd_dt).days,
                'Is_Aligned': (due_dt - prd_dt).days == 0,
                'Alignment_Status': 'Aligned' if (due_dt - prd_dt).days == 0 
                                else 'WO_Early' if (due_dt - prd_dt).days < 0 
                                else 'WO_Late'
            }
            data.append(record)
        
        self.so_wo_data = pd.DataFrame(data)
        
        # Calcular métricas
        total = len(self.so_wo_data)
        aligned = self.so_wo_data['Is_Aligned'].sum()
        wo_early = (self.so_wo_data['Date_Diff_Days'] < 0).sum()
        wo_late = (self.so_wo_data['Date_Diff_Days'] > 0).sum()
        
        self.so_wo_metrics = {
            'total_records': total,
            'aligned_count': int(aligned),
            'wo_early_count': int(wo_early),
            'wo_late_count': int(wo_late), 
            'alignment_pct': round(100 * aligned / total, 2),
            'mean_diff_days': round(self.so_wo_data['Date_Diff_Days'].mean(), 2),
            'std_diff_days': round(self.so_wo_data['Date_Diff_Days'].std(), 2),
            'median_diff_days': round(self.so_wo_data['Date_Diff_Days'].median(), 2),
            'so_total': total + 50,  # Simular que había más SO sin WO
            'wo_total': total,
            'match_rate': round(100 * total / (total + 50), 2)
        }

    def create_so_wo_alignment_chart(self):
        """Crear gráfico de alineación SO-WO"""
        if not hasattr(self, 'so_wo_data') or self.so_wo_data is None:
            return None
        
        try:
            import matplotlib
            matplotlib.use('Agg')
            
            plt.style.use('dark_background')
            fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 12))
            fig.patch.set_facecolor(COLORS['surface'])
            
            # GRÁFICO 1: Histograma de diferencias
            diff_data = self.so_wo_data['Date_Diff_Days'].dropna()
            
            # Filtrar outliers para mejor visualización
            q1, q3 = diff_data.quantile([0.25, 0.75])
            iqr = q3 - q1
            lower_bound = q1 - 2 * iqr
            upper_bound = q3 + 2 * iqr
            filtered_data = diff_data[(diff_data >= lower_bound) & (diff_data <= upper_bound)]
            
            # Histograma
            n, bins, patches = ax1.hist(filtered_data, bins=50, alpha=0.7, 
                                        density=True, edgecolor='white', linewidth=0.3)
            
            # Colorear barras: verde para alineados, rojo/azul para desalineados
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
            metrics = self.so_wo_metrics
            textstr = f'🟢 Alineados: {metrics["aligned_count"]:,} ({metrics["alignment_pct"]:.1f}%)\n🔴 WO Temprano: {metrics["wo_early_count"]:,}\n🔵 WO Tardío: {metrics["wo_late_count"]:,}\nTotal: {metrics["total_records"]:,}'
            props = dict(boxstyle='round', facecolor='black', alpha=0.8)
            ax1.text(0.02, 0.98, textstr, transform=ax1.transAxes, fontsize=10,
                    verticalalignment='top', bbox=props, color='white')
            
            ax1.set_title('SO vs WO: DueDt - Prd_Dt\n(0 = Perfecta Alineación)', 
                        color='white', fontsize=14, fontweight='bold')
            ax1.set_xlabel('Diferencia (días)', color='white', fontsize=12)
            ax1.set_ylabel('Densidad', color='white', fontsize=12)
            ax1.grid(True, alpha=0.3, color='white')
            ax1.set_facecolor('#2C3E50')
            ax1.tick_params(colors='white')
            for spine in ax1.spines.values():
                spine.set_color('white')
            
            # GRÁFICO 2: Distribución por categorías
            categories = ['WO Temprano\n(Antes SO)', 'Alineados\n(Exacto)', 'WO Tardío\n(Después SO)']
            values = [metrics['wo_early_count'], metrics['aligned_count'], metrics['wo_late_count']]
            colors = ['#FF6B6B', '#4ECDC4', '#45B7D1']
            
            bars = ax2.bar(categories, values, color=colors, alpha=0.8, edgecolor='white', linewidth=1)
            
            # Agregar valores en las barras
            total_records = metrics['total_records']
            for bar, value in zip(bars, values):
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height + total_records*0.01,
                        f'{value:,}\n({100*value/total_records:.1f}%)',
                        ha='center', va='bottom', color='white', fontsize=11, fontweight='bold')
            
            ax2.set_title('Distribución SO-WO Alignment', color='white', fontsize=14, fontweight='bold')
            ax2.set_ylabel('Cantidad', color='white', fontsize=12)
            ax2.set_facecolor('#2C3E50')
            ax2.tick_params(colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='y')
            for spine in ax2.spines.values():
                spine.set_color('white')
            
            # GRÁFICO 3: Análisis por AC (si hay datos)
            if 'AC_so' in self.so_wo_data.columns:
                ac_stats = self.so_wo_data.groupby('AC_so')['Is_Aligned'].agg(['count', 'sum']).reset_index()
                ac_stats['alignment_pct'] = 100 * ac_stats['sum'] / ac_stats['count']
                
                ax3.bar(ac_stats['AC_so'], ac_stats['alignment_pct'], 
                    color='#45B7D1', alpha=0.8, edgecolor='white')
                ax3.set_title('Alineación por AC', color='white', fontsize=14, fontweight='bold')
                ax3.set_xlabel('AC', color='white')
                ax3.set_ylabel('% Alineación', color='white')
                ax3.set_facecolor('#2C3E50')
                ax3.tick_params(colors='white')
                ax3.grid(True, alpha=0.3, color='white', axis='y')
                for spine in ax3.spines.values():
                    spine.set_color('white')
            else:
                ax3.text(0.5, 0.5, 'AC Analysis\nNo disponible', transform=ax3.transAxes,
                        ha='center', va='center', color='white', fontsize=12)
                ax3.set_facecolor('#2C3E50')
            
            # GRÁFICO 4: Análisis temporal (diferencias a lo largo del tiempo)
            if 'Prd_Dt' in self.so_wo_data.columns:
                # Crear bins por semanas
                self.so_wo_data['Week'] = self.so_wo_data['Prd_Dt'].dt.to_period('W')
                weekly_stats = self.so_wo_data.groupby('Week')['Date_Diff_Days'].agg(['mean', 'count']).reset_index()
                weekly_stats = weekly_stats[weekly_stats['count'] >= 5]  # Solo semanas con suficientes datos
                
                if len(weekly_stats) > 0:
                    ax4.plot(range(len(weekly_stats)), weekly_stats['mean'], 
                            marker='o', color='#4ECDC4', linewidth=2, markersize=6)
                    ax4.axhline(y=0, color='yellow', linestyle='-', linewidth=2, alpha=0.7)
                    ax4.set_title('Tendencia Temporal\n(Diferencia Promedio por Semana)', 
                                color='white', fontsize=14, fontweight='bold')
                    ax4.set_xlabel('Semanas', color='white')
                    ax4.set_ylabel('Días de Diferencia', color='white')
                    ax4.set_facecolor('#2C3E50')
                    ax4.tick_params(colors='white')
                    ax4.grid(True, alpha=0.3, color='white')
                    for spine in ax4.spines.values():
                        spine.set_color('white')
                else:
                    ax4.text(0.5, 0.5, 'Análisis Temporal\nDatos insuficientes', transform=ax4.transAxes,
                            ha='center', va='center', color='white', fontsize=12)
                    ax4.set_facecolor('#2C3E50')
            else:
                ax4.text(0.5, 0.5, 'Análisis Temporal\nNo disponible', transform=ax4.transAxes,
                        ha='center', va='center', color='white', fontsize=12)
                ax4.set_facecolor('#2C3E50')
            
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
        analyzer.load_so_wo_alignment_data()  # NUEVA LÍNEA
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
            # export_card = create_export_card(analyzer)
            export_card = create_export_card(analyzer, lambda _: export_data())
            
            kpi_row1 = ft.Row([
                create_metric_card(
                    "Precisión Exacta", 
                    f"{analyzer.metrics.get('exact_match_pct', 0)}%",
                    "Coincidencia perfecta",
                    COLORS['success']
                ),
                create_metric_card(
                    "Score de Calidad", 
                    f"{analyzer.metrics.get('quality_score', 0)}",
                    "Índice 0-100",
                    COLORS['accent']
                ),
                create_metric_card(
                    "Diferencia Media", 
                    f"{analyzer.metrics.get('mean_diff', 0)} días",
                    f"±{analyzer.metrics.get('std_diff', 0)} desv.",
                    COLORS['warning']
                ),
                create_metric_card(
                    "Entregas Puntuales", 
                    f"{analyzer.metrics.get('on_time_deliveries_pct', 0)}%",
                    "Ship = Req Date",
                    COLORS['success']
                )
            ], wrap=True, spacing=15)
            
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
                #kpi_row1,
                chart_container
            ], spacing=20)
        


        # Modifica la función create_advanced_metrics() para usar la validación FCST
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
            
            # Botón de exportación FCST
            def export_fcst_analysis():
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    export_path = f"FCST_Alignment_Analysis_{timestamp}.xlsx"
                    
                    with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                        # Datos completos
                        analyzer.fcst_data.to_excel(writer, sheet_name='FCST_Complete', index=False)
                        
                        # Solo desalineados
                        misaligned = analyzer.fcst_data[~analyzer.fcst_data['Is_Aligned']]
                        misaligned.to_excel(writer, sheet_name='FCST_Misaligned', index=False)
                        
                        # Solo alineados
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


    # Nueva función para el tercer tab
        def create_so_wo_metrics(analyzer, page):
                """Crear pestaña de métricas SO-WO"""
                # Cargar datos de SO-WO si no están cargados
                if not hasattr(analyzer, 'so_wo_data'):
                    analyzer.load_so_wo_alignment_data()
                
                # Métricas de SO-WO
                metrics = analyzer.so_wo_metrics
                
                # Tarjetas de métricas SO-WO
                so_wo_cards = ft.Row([
                    create_metric_card(
                        "SO con WO", 
                        f"{metrics.get('total_records', 0):,}",
                        f"{metrics.get('match_rate', 0)}% del total SO",
                        COLORS['accent']
                    ),
                    create_metric_card(
                        "Alineación SO-WO", 
                        f"{metrics.get('alignment_pct', 0)}%",
                        f"{metrics.get('aligned_count', 0):,} perfectos",
                        COLORS['success']
                    ),
                    create_metric_card(
                        "WO Temprano", 
                        f"{metrics.get('wo_early_count', 0):,}",
                        "WO vence antes SO",
                        COLORS['error']
                    ),
                    create_metric_card(
                        "WO Tardío", 
                        f"{metrics.get('wo_late_count', 0):,}",
                        "WO vence después SO",
                        COLORS['warning']
                    )
                ], wrap=True, spacing=15)
                
                # Gráfico de alineación SO-WO
                chart_base64 = analyzer.create_so_wo_alignment_chart()
                chart_container = ft.Container(
                    content=ft.Column([
                        ft.Text("📊 ANÁLISIS DE ALINEACIÓN SALES ORDER vs WORK ORDER", 
                            size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                        ft.Text("Comparación entre Prd_Dt (SO) y DueDt (WO) usando llave SO_No-Ln = SO_FCST-Sub", 
                            size=12, color=COLORS['text_secondary']),
                        ft.Image(
                            src_base64=chart_base64,
                            width=page.width - 80,
                            height=600,
                            fit=ft.ImageFit.CONTAIN
                        ) if chart_base64 else ft.Text("Error generando gráfico", color=COLORS['error'])
                    ]),
                    padding=ft.padding.all(20),
                    bgcolor=COLORS['primary'],
                    border_radius=12
                )
                
                # Botón de exportación SO-WO
                def export_so_wo_analysis():
                    try:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        export_path = f"SO_WO_Alignment_Analysis_{timestamp}.xlsx"
                        
                        # Columnas seleccionadas para exportar
                        selected_columns = [
                            # SO Columns
                            'Entity_so', 'SO_No', 'Ln', 'SO_Key', 'Item_Number', 
                            'Description_so', 'Prd_Dt', 'Opn_Q', 'Cust_Name', 'AC_so', 'WO_No',
                            # WO Columns  
                            'WONo', 'SO_FCST', 'Sub', 'DueDt', 'Description_wo', 
                            'WoType', 'AC_wo', 'St',
                            # Analysis
                            'Date_Diff_Days', 'Is_Aligned', 'Alignment_Status'
                        ]
                        
                        # Filtrar solo columnas que existen
                        available_columns = [col for col in selected_columns if col in analyzer.so_wo_data.columns]
                        
                        with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                            # Datos completos
                            analyzer.so_wo_data[available_columns].to_excel(writer, sheet_name='SO_WO_Complete', index=False)
                            
                            # Solo alineados
                            aligned = analyzer.so_wo_data[analyzer.so_wo_data['Is_Aligned']]
                            aligned[available_columns].to_excel(writer, sheet_name='SO_WO_Aligned', index=False)
                            
                            # Solo desalineados
                            misaligned = analyzer.so_wo_data[~analyzer.so_wo_data['Is_Aligned']]
                            misaligned[available_columns].to_excel(writer, sheet_name='SO_WO_Misaligned', index=False)
                            
                            # WO Temprano
                            wo_early = analyzer.so_wo_data[analyzer.so_wo_data['Date_Diff_Days'] < 0]
                            wo_early[available_columns].to_excel(writer, sheet_name='WO_Early', index=False)
                            
                            # WO Tardío
                            wo_late = analyzer.so_wo_data[analyzer.so_wo_data['Date_Diff_Days'] > 0]
                            wo_late[available_columns].to_excel(writer, sheet_name='WO_Late', index=False)
                        
                        # Abrir archivo
                        full_path = os.path.abspath(export_path)
                        import subprocess
                        import platform
                        
                        if platform.system() == 'Windows':
                            subprocess.call(('start', full_path), shell=True)
                        
                        return f"✅ Exportado SO-WO: {export_path}"
                        
                    except Exception as e:
                        return f"❌ Error SO-WO: {str(e)}"
                
                export_so_wo_card = ft.Card(
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
                                    on_click=lambda _: print(export_so_wo_analysis()),
                                    bgcolor=COLORS['success'],
                                    color=COLORS['text_primary']
                                )
                            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                            ft.Text("Se exportarán: Completos, Alineados, Desalineados, WO Temprano, WO Tardío", 
                                    size=12, color=COLORS['text_secondary']),
                            ft.Divider(color=COLORS['secondary']),
                            ft.Row([
                                ft.Column([
                                    ft.Text("🟢 Alineados", size=12, color=COLORS['success']),
                                    ft.Text(f"{metrics.get('aligned_count', 0):,}", size=16, color=COLORS['success'], weight=ft.FontWeight.BOLD)
                                ], expand=1),
                                ft.Column([
                                    ft.Text("🔴 WO Temprano", size=12, color=COLORS['error']),
                                    ft.Text(f"{metrics.get('wo_early_count', 0):,}", size=16, color=COLORS['error'], weight=ft.FontWeight.BOLD)
                                ], expand=1),
                                ft.Column([
                                    ft.Text("🔵 WO Tardío", size=12, color=COLORS['warning']),
                                    ft.Text(f"{metrics.get('wo_late_count', 0):,}", size=16, color=COLORS['warning'], weight=ft.FontWeight.BOLD)
                                ], expand=1)
                            ])
                        ])
                    ),
                    elevation=8
                )
                
                return ft.Container(
                    content=ft.Column([
                        ft.Text("📈 ANÁLISIS DE ALINEACIÓN SALES ORDER vs WORK ORDER", 
                            size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                        ft.Text("Validación de fechas SO (Prd_Dt) vs WO (DueDt) con llave compuesta", 
                            size=14, color=COLORS['text_secondary']),
                        so_wo_cards,
                        export_so_wo_card,
                        chart_container
                    ], spacing=20),
                    padding=ft.padding.all(20)
                )


















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
                content=create_advanced_metrics(analyzer, page)  # PASAR PARÁMETROS
            ),
            ft.Tab(
                text="🛒 Análisis SO-WO",
                content=create_so_wo_metrics(analyzer, page)  # PASAR PARÁMETROS
            )
        ],
        expand=1
    )
        
        page.update()
    
    # Cargar datos iniciales
    load_data()
    
    # Agregar contenedor principal
    page.add(main_container)

if __name__ == "__main__":
    ft.app(target=main)