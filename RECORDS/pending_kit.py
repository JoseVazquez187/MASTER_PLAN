import flet as ft
import pandas as pd
import sqlite3
from datetime import datetime
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

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
        
        # Debug inicial
        print(f"🔍 Sort codes únicos encontrados: {sorted(df['Srt'].unique())}")
        
        # Filtrar por sort codes ANTES de otros filtros
        df_sin_excluidos = df[~df['Srt'].isin(sort_codes_excluidos)].copy()
        excluidos_count = len(df) - len(df_sin_excluidos)
        print(f"🚫 Registros excluidos por sort codes: {excluidos_count}")
        print(f"📋 Registros después de excluir sort codes: {len(df_sin_excluidos)}")
        
        # Debug: Contar registros con QtyPending > 0
        pending_count = len(df_sin_excluidos[df_sin_excluidos['QtyPending'] > 0])
        print(f"📋 Registros con QtyPending > 0 (sin excluidos): {pending_count}")
        
        # Fecha actual
        hoy = pd.Timestamp.today()
        
        # Filtrar past due o del mes de junio
        df_filtrado = df_sin_excluidos[
            (df_sin_excluidos['ReqDate'] < hoy) | 
            (df_sin_excluidos['ReqDate'].dt.month == 6)
        ].copy()
        print(f"📅 Después de filtro de fechas: {len(df_filtrado)}")
        
        # *** FILTRO PRINCIPAL: Solo considerar líneas con QtyPending > 0 ***
        df_filtrado = df_filtrado[df_filtrado['QtyPending'] > 0].copy()
        print(f"🎯 Solo líneas pendientes (QtyPending > 0): {len(df_filtrado)}")
        
        if len(df_filtrado) == 0:
            print("⚠️ No hay líneas pendientes después de aplicar filtros")
            self.df_filtered = df_filtrado
            return True
        
        # Debug: Mostrar sort codes en datos finales
        sort_codes_finales = sorted(df_filtrado['Srt'].unique())
        print(f"✅ Sort codes en datos finales: {sort_codes_finales}")
        
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
            inventario_disponible = float(grupo['QtyOh'].iloc[0])  # Todos deben tener el mismo QtyOh
            print(f"🔧 Procesando {component}: OH={inventario_disponible}, Demandas={len(grupo)}")
            
            # Procesar demandas secuencialmente (ya ordenadas por fecha)
            for idx, row in grupo.iterrows():
                qty_pending = float(row['QtyPending'])
                std_cost = float(row['STDCost'])
                
                if inventario_disponible >= qty_pending:
                    # ✅ Tenemos suficiente inventario para cubrir COMPLETAMENTE esta demanda
                    df_filtrado.at[idx, 'COB'] = 'Sí'
                    df_filtrado.at[idx, 'Valor_Cubierto'] = qty_pending * std_cost     # Toda la demanda cubierta
                    df_filtrado.at[idx, 'Valor_No_Cubierto'] = 0                      # Nada sin cubrir
                    df_filtrado.at[idx, 'Qty_Faltante'] = 0                           # No falta nada
                    
                    # Consumir inventario
                    inventario_disponible -= qty_pending
                    df_filtrado.at[idx, 'Balance'] = inventario_disponible            # Inventario restante
                    
                elif inventario_disponible > 0:
                    # 🟡 Tenemos inventario PARCIAL para esta demanda
                    df_filtrado.at[idx, 'COB'] = 'Parcial'
                    qty_cubierta = inventario_disponible                              # Solo podemos cubrir lo que tenemos
                    qty_faltante = qty_pending - inventario_disponible               # Lo que nos falta
                    
                    df_filtrado.at[idx, 'Valor_Cubierto'] = qty_cubierta * std_cost
                    df_filtrado.at[idx, 'Valor_No_Cubierto'] = qty_faltante * std_cost
                    df_filtrado.at[idx, 'Qty_Faltante'] = qty_faltante
                    
                    # Inventario se agota completamente
                    inventario_disponible = 0
                    df_filtrado.at[idx, 'Balance'] = 0
                    
                else:
                    # ❌ No tenemos inventario para esta demanda
                    df_filtrado.at[idx, 'COB'] = 'No'
                    df_filtrado.at[idx, 'Valor_Cubierto'] = 0                         # Nada cubierto
                    df_filtrado.at[idx, 'Valor_No_Cubierto'] = qty_pending * std_cost # Todo sin cubrir
                    df_filtrado.at[idx, 'Qty_Faltante'] = qty_pending                 # Todo falta
                    df_filtrado.at[idx, 'Balance'] = 0                                # Sin inventario
        
        print(f"✅ Procesamiento secuencial completado: {len(df_filtrado)} registros analizados")
        self.df_filtered = df_filtrado
        return True
    
    def get_summary_metrics(self):
        """Calcular métricas resumen para el dashboard"""
        if self.df_filtered is None or len(self.df_filtered) == 0:
            return {
                'total_lines': 0,
                'covered_lines': 0,
                'partial_lines': 0,
                'not_covered_lines': 0,
                'coverage_percentage': 0,
                'total_value_required': 0,
                'total_value_not_covered': 0,
                'total_value_covered': 0,
                'total_qty_pending': 0,
                'wo_clear_count': 0,
                'wo_clear_value': 0,
                'wo_total_count': 0,
                'wo_clear_percentage': 0
            }
        
        total_lines = len(self.df_filtered)
        covered_lines = len(self.df_filtered[self.df_filtered['COB'] == 'Sí'])
        partial_lines = len(self.df_filtered[self.df_filtered['COB'] == 'Parcial'])
        not_covered_lines = len(self.df_filtered[self.df_filtered['COB'] == 'No'])
        
        # Calcular valores basados SOLO en cantidades pendientes
        total_value_required = (self.df_filtered['QtyPending'] * self.df_filtered['STDCost']).sum()
        total_value_not_covered = self.df_filtered['Valor_No_Cubierto'].sum()
        total_value_covered = self.df_filtered['Valor_Cubierto'].sum()
        total_qty_pending = self.df_filtered['QtyPending'].sum()
        
        coverage_percentage = (covered_lines / total_lines * 100) if total_lines > 0 else 0
        
        # *** NUEVO: Análisis de Work Orders completamente cubiertas ***
        wo_analysis = self._analyze_wo_coverage()
        
        return {
            'total_lines': total_lines,
            'covered_lines': covered_lines,
            'partial_lines': partial_lines,
            'not_covered_lines': not_covered_lines,
            'coverage_percentage': coverage_percentage,
            'total_value_required': total_value_required,
            'total_value_not_covered': total_value_not_covered,
            'total_value_covered': total_value_covered,
            'total_qty_pending': total_qty_pending,
            'wo_clear_count': wo_analysis['wo_clear_count'],
            'wo_clear_value': wo_analysis['wo_clear_value'],
            'wo_total_count': wo_analysis['wo_total_count'],
            'wo_clear_percentage': wo_analysis['wo_clear_percentage']
        }
    
    def _analyze_wo_coverage(self):
        """Analizar cobertura por Work Order"""
        if self.df_filtered is None or len(self.df_filtered) == 0:
            return {
                'wo_clear_count': 0,
                'wo_clear_value': 0,
                'wo_total_count': 0,
                'wo_clear_percentage': 0
            }
        
        # Agrupar por Work Order (columna WoNo)
        wo_groups = self.df_filtered.groupby('WoNo')
        
        wo_clear_count = 0
        wo_clear_value = 0
        wo_total_count = len(wo_groups)
        
        for wo_number, wo_data in wo_groups:
            # Verificar si TODOS los componentes de esta WO están completamente cubiertos
            all_covered = all(wo_data['COB'] == 'Sí')
            
            if all_covered:
                wo_clear_count += 1
                # Sumar el valor total de esta WO (suma de todos sus componentes pendientes)
                wo_value = (wo_data['QtyPending'] * wo_data['STDCost']).sum()
                wo_clear_value += wo_value
                
                print(f"✅ WO {wo_number}: Completamente cubierta - Valor: ${wo_value:,.2f}")
            else:
                # Debug: mostrar componentes no cubiertos
                not_covered = wo_data[wo_data['COB'] != 'Sí']
                print(f"❌ WO {wo_number}: {len(not_covered)} componentes sin cobertura completa")
        
        wo_clear_percentage = (wo_clear_count / wo_total_count * 100) if wo_total_count > 0 else 0
        
        print(f"📊 Resumen WO: {wo_clear_count}/{wo_total_count} WOs completamente cubiertas ({wo_clear_percentage:.1f}%)")
        print(f"💰 Valor total WO Clear: ${wo_clear_value:,.2f}")
        
        return {
            'wo_clear_count': wo_clear_count,
            'wo_clear_value': wo_clear_value,
            'wo_total_count': wo_total_count,
            'wo_clear_percentage': wo_clear_percentage
        }
    
    def export_to_excel(self, export_path=None):
        """Exportar análisis completo a Excel"""
        if self.df_filtered is None or len(self.df_filtered) == 0:
            return {"success": False, "message": "No hay datos pendientes para exportar"}
        
        try:
            if export_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                export_path = f"Analisis_Cobertura_Pendientes_{timestamp}.xlsx"
            
            # Separar datos por categorías
            datos_cubiertos = self.df_filtered[self.df_filtered['COB'] == 'Sí'].copy()
            datos_no_cubiertos = self.df_filtered[self.df_filtered['COB'] == 'No'].copy()
            
            # Calcular métricas resumen
            metrics = self.get_summary_metrics()
            
            # Crear DataFrame de resumen
            resumen_data = {
                'Métrica': [
                    'Total de Líneas Pendientes',
                    'Líneas Pendientes Cubiertas por OH',
                    'Líneas Pendientes Parcialmente Cubiertas',
                    'Líneas Pendientes NO Cubiertas por OH',
                    '% Cobertura por OH',
                    'Cantidad Total Pendiente (unidades)',
                    'Valor Total Pendiente',
                    'Valor Pendiente Cubierto por OH',
                    'Valor Pendiente NO Cubierto por OH',
                    '% Valor Pendiente Cubierto por OH',
                    'Work Orders Completamente Cubiertas',
                    'Valor de WOs Completamente Cubiertas',
                    'Total Work Orders Analizadas',
                    '% WOs Completamente Cubiertas'
                ],
                'Valor': [
                    f"{metrics['total_lines']:,}",
                    f"{metrics['covered_lines']:,}",
                    f"{metrics['partial_lines']:,}",
                    f"{metrics['not_covered_lines']:,}",
                    f"{metrics['coverage_percentage']:.2f}%",
                    f"{metrics['total_qty_pending']:,.0f}",
                    f"${metrics['total_value_required']:,.2f}",
                    f"${metrics['total_value_covered']:,.2f}",
                    f"${metrics['total_value_not_covered']:,.2f}",
                    f"{(metrics['total_value_covered']/metrics['total_value_required']*100) if metrics['total_value_required'] > 0 else 0:.2f}%",
                    f"{metrics['wo_clear_count']:,}",
                    f"${metrics['wo_clear_value']:,.2f}",
                    f"{metrics['wo_total_count']:,}",
                    f"{metrics['wo_clear_percentage']:.2f}%"
                ]
            }
            resumen_df = pd.DataFrame(resumen_data)
            
            # Exportar a Excel con múltiples hojas
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                # Hoja de resumen ejecutivo
                resumen_df.to_excel(writer, sheet_name='Resumen_Ejecutivo', index=False)
                
                # Hoja con todos los datos pendientes
                self.df_filtered.to_excel(writer, sheet_name='Datos_Pendientes', index=False)
                
                # Hoja con líneas pendientes cubiertas por OH
                if len(datos_cubiertos) > 0:
                    datos_cubiertos.to_excel(writer, sheet_name='Pendientes_Cubiertos_OH', index=False)
                else:
                    empty_df = pd.DataFrame(columns=self.df_filtered.columns)
                    empty_df.to_excel(writer, sheet_name='Pendientes_Cubiertos_OH', index=False)
                
                # Hoja con líneas pendientes NO cubiertas por OH
                if len(datos_no_cubiertos) > 0:
                    datos_no_cubiertos.to_excel(writer, sheet_name='Pendientes_NO_Cubiertos_OH', index=False)
                else:
                    empty_df = pd.DataFrame(columns=self.df_filtered.columns)
                    empty_df.to_excel(writer, sheet_name='Pendientes_NO_Cubiertos_OH', index=False)
            
            # Convertir a ruta absoluta
            import os
            full_path = os.path.abspath(export_path)
            
            return {
                "success": True,
                "message": f"Exportado: {os.path.basename(export_path)}",
                "path": full_path,
                "covered_lines": len(datos_cubiertos),
                "not_covered_lines": len(datos_no_cubiertos),
                "total_lines": len(self.df_filtered)
            }
            
        except Exception as e:
            print(f"ERROR en exportación: {e}")
            return {"success": False, "message": f"Error: {str(e)}"}

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
    # Referencias para actualizar estado
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
            result = dashboard.export_to_excel()
            
            if result["success"]:
                status_text_ref.current.value = f"✅ {result['message']} - Abriendo archivo..."
                status_text_ref.current.color = COLORS['success']
                export_button_ref.current.text = "✅ Completado"
                
                # Abrir el archivo automáticamente
                try:
                    import subprocess
                    import platform
                    import os
                    
                    file_path = result['path']
                    
                    if platform.system() == 'Darwin':       # macOS
                        subprocess.call(('open', file_path))
                    elif platform.system() == 'Windows':   # Windows
                        os.startfile(file_path)
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
            export_button_ref.current.text = "📊 Exportar a Excel"
            export_button_ref.current.disabled = False
            export_button_ref.current.update()
        
        threading.Thread(target=restore_button, daemon=True).start()
    
    # Obtener métricas para mostrar preview
    metrics = dashboard.get_summary_metrics() if dashboard.df_filtered is not None else {}
    
    return ft.Card(
        content=ft.Container(
            padding=ft.padding.all(20),
            bgcolor=COLORS['primary'],
            border_radius=12,
            content=ft.Column([
                ft.Row([
                    ft.Text("📤 Exportar Análisis de Pendientes", 
                            size=18, color=COLORS['text_primary'], weight=ft.FontWeight.BOLD),
                    ft.ElevatedButton(
                        "📊 Exportar a Excel",
                        on_click=handle_export,
                        bgcolor=COLORS['success'],
                        color=COLORS['text_primary'],
                        ref=export_button_ref,
                        icon=ft.Icons.FILE_DOWNLOAD
                    )
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                
                # Área de estado
                ft.Text(
                    "Listo para exportar líneas pendientes",
                    size=12,
                    color=COLORS['text_secondary'],
                    ref=status_text_ref
                ),
                
                ft.Divider(color=COLORS['secondary']),
                ft.Text("📋 Se exportarán las siguientes hojas:", size=12, color=COLORS['text_secondary']),
                ft.Column([
                    ft.Text("• Resumen_Ejecutivo (métricas principales)", size=11, color=COLORS['accent']),
                    ft.Text("• Datos_Pendientes (solo líneas con QtyPending > 0)", size=11, color=COLORS['text_secondary']),
                    ft.Text("• Pendientes_Cubiertos_OH (pendiente cubierto por OH)", size=11, color=COLORS['success']),
                    ft.Text("• Pendientes_NO_Cubiertos_OH (pendiente NO cubierto)", size=11, color=COLORS['error']),
                ], spacing=2)
            ], spacing=10)
        ),
        elevation=8
    )

def main(page: ft.Page):
    page.title = "Dashboard Ejecutivo - Cobertura de Inventario"
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = COLORS['surface']
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20
    
    # Instancia del dashboard
    dashboard = InventoryDashboard()
    
    # Contenedor principal
    main_container = ft.Column(spacing=20)
    
    # Estado para notificaciones
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
    
    # Función para actualizar el dashboard
    def update_dashboard():
        # Limpiar contenidos anteriores
        main_container.controls.clear()
        data_table.rows.clear()
        
        # Cargar y procesar datos
        if dashboard.load_data_from_db():
            if dashboard.process_data():
                # Obtener métricas
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
                for _, row in dashboard.df_filtered.head(100).iterrows():  # Limitar a 100 filas
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
                
                # Crear tarjeta de exportación
                export_card = create_export_card(dashboard, page)
                
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
    
    # Cargar datos iniciales
    update_dashboard()
    
    # Agregar contenedor principal
    page.add(main_container)

if __name__ == "__main__":
    ft.app(target=main)