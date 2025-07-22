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
        
    def load_data_from_db(self, db_path=r"C:\Users\J.Vazquez\Desktop\PARCHE EXPEDITE\BS_Analisis_2025_07_22.db"):
        """Cargar datos desde la base de datos SQLite"""
        try:
            print(f"Intentando conectar a: {db_path}")
            conn = sqlite3.connect(db_path)
            query = """
            SELECT EntityGroup, ItemNo, DemandSource, ShipDate, OH, MLIKCode, 
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
            "DemandSource": np.random.choice(demand_sources, n_records),
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
        """Exportar análisis CORREGIDO"""
        if self.df is None or len(self.df) == 0:
            return {"success": False, "message": "No hay datos para exportar"}
        
        try:
            # CORREGIDO: Exportar problemas (valores negativos) en lugar de entregas tardías
            backschedule_problems = self.df[self.df["Diff_Ship_Req"] < 0].copy()      # Material llega DESPUÉS (PROBLEMAS)
            original_problems = self.df[self.df["Diff_Ship_Original"] < 0].copy()     # Material llega DESPUÉS (PROBLEMAS)
            
            if export_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                export_path = f"Analisis_BackSchedule_Problemas_{timestamp}.xlsx"
            
            bs_problems_count = len(backschedule_problems)
            orig_problems_count = len(original_problems)
            
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                if len(backschedule_problems) > 0:
                    backschedule_problems.to_excel(writer, sheet_name='Problemas_BackSchedule', index=False)
                
                if len(original_problems) > 0:
                    original_problems.to_excel(writer, sheet_name='Problemas_Original', index=False)
                
                comparison_data = {
                    'Métrica': [
                        'Total Registros',
                        'BackSchedule - Problemas (material llega después)',
                        'Original - Problemas (material llega después)',
                        'Sistema Recomendado (Menos problemas)'
                    ],
                    'Valor': [
                        f"{len(self.df):,}",
                        f"{bs_problems_count:,}",
                        f"{orig_problems_count:,}",
                        "BackSchedule" if bs_problems_count < orig_problems_count else "Original"
                    ]
                }
                
                comparison_df = pd.DataFrame(comparison_data)
                comparison_df.to_excel(writer, sheet_name='Resumen_Final', index=False)
            
            return {
                "success": True,
                "message": f"Exportado: {export_path}",
                "path": export_path,
                "backschedule_problems": bs_problems_count,
                "original_problems": orig_problems_count,
                "better_system": "BackSchedule" if bs_problems_count < orig_problems_count else "Original"
            }
            
        except Exception as e:
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
    
    orig_data = analyzer.df["Diff_Ship_Original"].dropna()  # ShipDate - ReqDate_Original
    orig_problems = len(orig_data[orig_data < 0])           # Material llega DESPUÉS (PROBLEMAS)
    orig_good = len(orig_data[orig_data > 0])               # Material llega ANTES (BUENO - BackSchedule correcto)
    orig_ontime = len(orig_data[orig_data == 0])            # Material llega JUSTO A TIEMPO
    orig_total = len(orig_data)
    orig_leadtime_avg = orig_data[orig_data > 0].mean() if orig_good > 0 else 0
    
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
                        ft.Text(f"{100*bs_ontime/bs_total:.1f}%", size=10, color=COLORS['text_secondary'])
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
                        ft.Text(f"{100*orig_ontime/orig_total:.1f}%", size=10, color=COLORS['text_secondary'])
                    ], expand=2)
                ]),
                ft.Divider(color=COLORS['secondary']),
                ft.Text(f"BackSchedule promedio: {orig_leadtime_avg:.1f} días", size=12, color=COLORS['success'])
            ], spacing=10)
        ),
        elevation=8
    )
    
    return ft.Row([bs_card, orig_card], spacing=20)

def create_comparison_card(analyzer):
    """Crear tarjeta comparativa CORREGIDA"""
    if analyzer.df is None or len(analyzer.df) == 0:
        return ft.Card()
    
    summary = analyzer.get_effectiveness_summary()
    better_system = summary.get('better_system', 'N/A')
    better_color = COLORS['accent'] if better_system == "BackSchedule" else COLORS['success']
    
    # Estadísticas para BackSchedule promedio (solo valores positivos = material llega antes)
    bs_leadtime_avg = summary.get('bs_leadtime_avg', 0)
    orig_leadtime_avg = summary.get('orig_leadtime_avg', 0)
    
    return ft.Card(
        content=ft.Container(
            padding=ft.padding.all(25),
            bgcolor=COLORS['secondary'],
            border_radius=12,
            content=ft.Column([
                ft.Text("⚔️ COMPARATIVA: BackSchedule vs Original", 
                    size=20, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD,
                    text_align=ft.TextAlign.CENTER),
                ft.Text("🔵 = ShipDate - ReqDate | 🟢 = ShipDate - ReqDate_Original", 
                    size=12, color=COLORS['text_secondary'], text_align=ft.TextAlign.CENTER),
                ft.Divider(color=COLORS['text_secondary']),
                
                ft.Row([
                    # Columna de PROBLEMAS (valores negativos = material llega después)
                    ft.Column([
                        ft.Text("❌ PROBLEMAS DE ABASTECIMIENTO", size=14, color=COLORS['error'], weight=ft.FontWeight.BOLD),
                        ft.Text("(Material llega DESPUÉS de necesitarlo)", size=10, color=COLORS['text_secondary']),
                        ft.Container(height=10),
                        ft.Row([
                            ft.Container(
                                content=ft.Text("🔵", size=20),
                                width=30
                            ),
                            ft.Column([
                                ft.Text("BackSchedule:", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"{summary.get('backschedule_problems', 0):,} problemas", 
                                    size=14, color=COLORS['error'], weight=ft.FontWeight.BOLD)
                            ], spacing=2)
                        ]),
                        ft.Row([
                            ft.Container(
                                content=ft.Text("🟢", size=20),
                                width=30
                            ),
                            ft.Column([
                                ft.Text("Original:", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"{summary.get('original_problems', 0):,} problemas", 
                                        size=14, color=COLORS['error'], weight=ft.FontWeight.BOLD)
                            ], spacing=2)
                        ])
                    ], expand=1),
                    
                    ft.VerticalDivider(color=COLORS['text_secondary']),
                    
                    # Columna de BACKSCHEDULE CORRECTO (valores positivos = material llega antes)
                    ft.Column([
                        ft.Text("✅ BACKSCHEDULE CORRECTO", size=14, color=COLORS['success'], weight=ft.FontWeight.BOLD),
                        ft.Text("(Material llega ANTES de necesitarlo)", size=10, color=COLORS['text_secondary']),
                        ft.Container(height=10),
                        ft.Row([
                            ft.Container(
                                content=ft.Text("🔵", size=20),
                                width=30
                            ),
                            ft.Column([
                                ft.Text("BackSchedule:", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"{summary.get('backschedule_good', 0):,} casos", 
                                        size=14, color=COLORS['success'], weight=ft.FontWeight.BOLD)
                            ], spacing=2)
                        ]),
                        ft.Row([
                            ft.Container(
                                content=ft.Text("🟢", size=20),
                                width=30
                            ),
                            ft.Column([
                                ft.Text("Original:", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"{summary.get('original_good', 0):,} casos", 
                                        size=14, color=COLORS['success'], weight=ft.FontWeight.BOLD)
                            ], spacing=2)
                        ])
                    ], expand=1),
                    
                    ft.VerticalDivider(color=COLORS['text_secondary']),
                    
                    # Columna de JUSTO A TIEMPO (valores cero = material llega exacto)
                    ft.Column([
                        ft.Text("⭕ JUSTO A TIEMPO", size=14, color=COLORS['warning'], weight=ft.FontWeight.BOLD),
                        ft.Text("(Material llega EXACTO)", size=10, color=COLORS['text_secondary']),
                        ft.Container(height=10),
                        ft.Row([
                            ft.Container(
                                content=ft.Text("🔵", size=20),
                                width=30
                            ),
                            ft.Column([
                                ft.Text("BackSchedule:", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"{summary.get('backschedule_ontime', 0):,} exactos", 
                                        size=14, color=COLORS['warning'], weight=ft.FontWeight.BOLD)
                            ], spacing=2)
                        ]),
                        ft.Row([
                            ft.Container(
                                content=ft.Text("🟢", size=20),
                                width=30
                            ),
                            ft.Column([
                                ft.Text("Original:", size=12, color=COLORS['text_secondary']),
                                ft.Text(f"{summary.get('original_ontime', 0):,} exactos", 
                                        size=14, color=COLORS['warning'], weight=ft.FontWeight.BOLD)
                            ], spacing=2)
                        ])
                    ], expand=1)
                ]),
                
                ft.Divider(color=COLORS['text_secondary']),
                
                # BackSchedule promedio para casos correctos
                ft.Container(
                    content=ft.Column([
                        ft.Text("📊 BACKSCHEDULE PROMEDIO (Solo casos con material antes):", size=14, color=COLORS['text_secondary'],
                                text_align=ft.TextAlign.CENTER),
                        ft.Row([
                            ft.Column([
                                ft.Text("🔵 BackSchedule", size=12, color=COLORS['accent']),
                                ft.Text(f"{bs_leadtime_avg:.1f} días antes", size=16, color=COLORS['success'], weight=ft.FontWeight.BOLD)
                            ], expand=1),
                            ft.Column([
                                ft.Text("🟢 Original", size=12, color=COLORS['success']),
                                ft.Text(f"{orig_leadtime_avg:.1f} días antes", size=16, color=COLORS['success'], weight=ft.FontWeight.BOLD)
                            ], expand=1)
                        ])
                    ], spacing=5),
                    bgcolor=COLORS['primary'],
                    padding=ft.padding.all(15),
                    border_radius=8
                ),
                
                # Veredicto final
                ft.Container(
                    content=ft.Column([
                        ft.Text("🏆 SISTEMA GANADOR (Menos problemas de abastecimiento):", size=16, color=COLORS['text_secondary'],
                                text_align=ft.TextAlign.CENTER),
                        ft.Text(better_system, size=24, color=better_color, 
                                weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER),
                        ft.Text(f"({abs(summary.get('difference', 0)):,} menos problemas)", 
                                size=12, color=COLORS['text_secondary'], text_align=ft.TextAlign.CENTER)
                    ], spacing=5),
                    bgcolor=COLORS['primary'],
                    padding=ft.padding.all(15),
                    border_radius=8
                ),
                
                # Explicación
                ft.Container(
                    content=ft.Column([
                        ft.Text("💡 INTERPRETACIÓN CORRECTA:", size=14, color=COLORS['warning'], weight=ft.FontWeight.BOLD),
                        ft.Text("Positivo (+) = Material llega ANTES ✅ | Cero (0) = Material llega EXACTO ⭕ | Negativo (-) = Material llega DESPUÉS ❌", 
                                size=11, color=COLORS['text_secondary'], text_align=ft.TextAlign.CENTER),
                        ft.Text("Menos problemas de abastecimiento = Mejor sistema de planificación", 
                                size=12, color=COLORS['text_secondary'], text_align=ft.TextAlign.CENTER)
                    ], spacing=3),
                    bgcolor=COLORS['primary'],
                    padding=ft.padding.all(10),
                    border_radius=8
                )
            ], spacing=15)
        ),
        elevation=12
    )

def create_export_card(analyzer):
    """Crear tarjeta de exportación CORREGIDA"""
    summary = analyzer.get_effectiveness_summary()
    
    return ft.Card(
        content=ft.Container(
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12,
            content=ft.Column([
                ft.Text("📤 Exportar Análisis Final", 
                        size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
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
                ft.Divider(color=COLORS['secondary']),
                ft.Text(f"🏆 Mejor: {summary.get('better_system', 'N/A')}", 
                        size=14, color=COLORS['warning'], weight=ft.FontWeight.BOLD),
                ft.Text("(Basado en menor cantidad de problemas de abastecimiento)", 
                        size=10, color=COLORS['text_secondary'])
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
        """Actualizar dashboard"""
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
                        ft.ElevatedButton(
                            "📊 Exportar",
                            on_click=lambda _: export_data(),
                            bgcolor=COLORS['error'],
                            color=COLORS['text_primary']
                        )
                    ], spacing=10),
                    export_status
                ], horizontal_alignment=ft.CrossAxisAlignment.END)
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
        
        # Tarjetas
        gaussian_cards = create_gaussian_cards(analyzer)
        comparison_card = create_comparison_card(analyzer)
        export_card = create_export_card(analyzer)
        
        # KPIs
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
        
        # Gráfico
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
        
        # Tabla
        table_data = analyzer.df.head(20) if analyzer.df is not None else pd.DataFrame()
        
        table_rows = []
        for _, row in table_data.iterrows():
            bs_diff = row.get('Diff_Ship_Req', 0)        # ShipDate - ReqDate
            orig_diff = row.get('Diff_Ship_Original', 0) # ShipDate - ReqDate_Original
            
            # CORREGIDO: Colores según nueva interpretación FINAL
            bs_color = COLORS['success'] if bs_diff > 0 else COLORS['error'] if bs_diff < 0 else COLORS['warning']
            orig_color = COLORS['success'] if orig_diff > 0 else COLORS['error'] if orig_diff < 0 else COLORS['warning']
                
            table_rows.append(
                ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(row.get('ItemNo', '')), color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(str(row.get('EntityGroup', '')), color=COLORS['text_secondary'])),
                    ft.DataCell(ft.Text(row.get('ShipDate', '').strftime("%Y-%m-%d") 
                                        if pd.notna(row.get('ShipDate')) else '', color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(row.get('ReqDate', '').strftime("%Y-%m-%d") 
                                        if pd.notna(row.get('ReqDate')) else '', color=COLORS['text_primary'])),
                    ft.DataCell(ft.Text(str(bs_diff), color=bs_color, weight=ft.FontWeight.BOLD)),
                    ft.DataCell(ft.Text(str(orig_diff), color=orig_color, weight=ft.FontWeight.BOLD))
                ])
            )
        
        table = ft.DataTable(
            bgcolor=COLORS['primary'],
            border=ft.border.all(1, COLORS['secondary']),
            border_radius=8,
            columns=[
                ft.DataColumn(label=ft.Text("Item", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(label=ft.Text("Entidad", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(label=ft.Text("Ship Date", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(label=ft.Text("Req Date", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(label=ft.Text("🔵 BS: Ship-Req", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD)),
                ft.DataColumn(label=ft.Text("🟢 Orig: Ship-Orig", color=COLORS['text_primary'], weight=ft.FontWeight.BOLD))
            ],
            rows=table_rows
        )
        
        table_container = ft.Container(
            content=ft.Column([
                ft.Text("Comparativo Final: Gestión de Abastecimiento (Top 20)", 
                        size=20, weight=ft.FontWeight.BOLD, color=COLORS['text_primary']),
                ft.Text("✅ Positivo = Material llega ANTES | ⭕ Cero = Material llega EXACTO | ❌ Negativo = Material llega DESPUÉS", 
                        size=12, color=COLORS['text_secondary']),
                ft.Container(content=table, border_radius=8)
            ]),
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12
        )
        
        # Agregar componentes
        main_container.controls.extend([
            header,
            gaussian_cards,
            comparison_card,
            export_card,
            kpi_row1,
            chart_container,
            table_container
        ])
        
        page.update()
    
    # Cargar datos iniciales
    load_data()
    
    # Agregar contenedor principal
    page.add(main_container)

if __name__ == "__main__":
    ft.app(target=main)