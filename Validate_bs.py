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
        
    def load_data_from_db(self, db_path=r"C:\Users\J.Vazquez\Desktop\PARCHE EXPEDITE\BS_Analisis_2025_07_30.db"):
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
        
        # Crear contenido de la segunda pestaña (Métricas avanzadas)
        def create_advanced_metrics():
            return ft.Container(
                content=ft.Column([
                    ft.Text("📈 MÉTRICAS AVANZADAS", 
                        size=24, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                    ft.Text("Análisis detallado y métricas adicionales", 
                        size=14, color=COLORS['text_secondary']),
                    ft.Container(
                        content=ft.Text("🚧 Contenido de métricas avanzadas en desarrollo...", 
                                    size=16, color=COLORS['warning']),
                        padding=ft.padding.all(40),
                        bgcolor=COLORS['primary'],
                        border_radius=12
                    )
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
                    text="📈 Métricas Avanzadas", 
                    content=create_advanced_metrics()
                )
            ],
            expand=1
        )
        
        # Agregar componentes
        main_container.controls.extend([
            header,
            tabs
        ])
        
        page.update()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    # Cargar datos iniciales
    load_data()
    
    # Agregar contenedor principal
    page.add(main_container)

if __name__ == "__main__":
    ft.app(target=main)