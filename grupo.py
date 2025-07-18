import flet as ft
import sqlite3
import pandas as pd
import os
from datetime import datetime, timedelta
import threading
import time

class KiteoDashboard:
    def __init__(self):
        self.theme_color = ft.Colors.PURPLE_700
        self.db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        self.desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.current_data = None
        self.is_processing = False
        
    def main(self, page: ft.Page):
        self.page = page
        page.title = "GRUPOS KITEO - Sistema de Propuesta"
        page.theme_mode = ft.ThemeMode.DARK
        page.padding = 20
        page.window_width = 1400
        page.window_height = 900
        page.window_center = True
        
        # Inicializar la interfaz
        self.show_dashboard()
    
    def show_snackbar(self, message, color=ft.Colors.BLUE_700):
        """Mostrar mensaje en snackbar"""
        snack_bar = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE),
            bgcolor=color
        )
        self.page.snack_bar = snack_bar
        snack_bar.open = True
        self.page.update()
    
    def process_kiteo_data(self, e):
        """Procesar datos de kiteo en hilo separado"""
        if self.is_processing:
            self.show_snackbar("⚠️ Procesamiento en curso...", ft.Colors.ORANGE_600)
            return
        
        def run_processing():
            try:
                self.is_processing = True
                self.update_status("🚀 Iniciando procesamiento...")
                
                # Verificar conexión a base de datos
                if not os.path.exists(self.db_path):
                    self.update_status("❌ Base de datos no encontrada")
                    self.show_snackbar("❌ No se puede conectar a la base de datos", ft.Colors.RED_600)
                    return
                
                self.update_status("📊 Conectando a la base de datos...")
                conn = sqlite3.connect(self.db_path)
                
                query = """
                SELECT 
                    Entity,
                    ItemNo,
                    Description,
                    WO,
                    ReqDate,
                    QtyFcst,
                    OpenQty,
                    Rev,
                    UM,
                    Proj
                FROM fcst 
                WHERE OpenQty > 0 
                AND WO IS NOT NULL 
                AND WO != ''
                ORDER BY ItemNo, Entity
                """
                
                df = pd.read_sql_query(query, conn)
                conn.close()
                
                self.update_status(f"✅ Datos cargados: {len(df)} registros")
                
                # Procesar fechas
                today = datetime.now().date()
                current_limit = today + timedelta(days=14)
                
                def simple_date_status(date_str):
                    try:
                        if pd.isna(date_str) or date_str == '' or date_str is None:
                            return "INVALID"
                        date_obj = pd.to_datetime(date_str, errors='coerce').date()
                        if pd.isna(date_obj):
                            return "INVALID"
                        if date_obj < today:
                            return "PAST_DUE"
                        elif date_obj <= current_limit:
                            return "CURRENT"
                        else:
                            return "FUTURE"
                    except:
                        return "INVALID"
                
                df['DateStatus'] = df['ReqDate'].apply(simple_date_status)
                
                self.update_status("📦 Generando propuesta: PAST_DUE y CURRENT+FUTURE agrupados")
                
                # PAST_DUE
                df_pastdue = df[df['DateStatus'] == 'PAST_DUE'].copy()
                propuesta_pastdue = df_pastdue.groupby(['ItemNo', 'Entity', 'DateStatus']).agg({
                    'WO': list,
                    'Description': 'first',
                    'OpenQty': 'sum',
                    'QtyFcst': 'sum',
                    'ReqDate': ['min', 'max'],
                    'Rev': 'first',
                    'UM': 'first',
                    'Proj': 'first'
                }).reset_index()
                
                # CURRENT + FUTURE
                df_cf = df[df['DateStatus'].isin(['CURRENT', 'FUTURE'])].copy()
                df_cf['DateStatus'] = 'CURRENT+FUTURE'
                propuesta_cf = df_cf.groupby(['ItemNo', 'Entity', 'DateStatus']).agg({
                    'WO': list,
                    'Description': 'first',
                    'OpenQty': 'sum',
                    'QtyFcst': 'sum',
                    'ReqDate': ['min', 'max'],
                    'Rev': 'first',
                    'UM': 'first',
                    'Proj': 'first'
                }).reset_index()
                
                # Combinar ambos bloques
                propuesta_final = pd.concat([propuesta_pastdue, propuesta_cf], ignore_index=True)
                propuesta_final.columns = ['ItemNo', 'Entity', 'DateStatus', 'WO_List', 'Description',
                                         'OpenQty_Total', 'QtyFcst_Total', 'ReqDate_Min', 'ReqDate_Max',
                                         'Rev', 'UM', 'Proj']
                
                propuesta_final['WO_Count'] = propuesta_final['WO_List'].apply(len)
                propuesta_final['WO_String'] = propuesta_final['WO_List'].apply(lambda x: ', '.join(map(str, x)))
                propuesta_final = propuesta_final.reset_index(drop=True)
                propuesta_final['Group_ID'] = range(1, len(propuesta_final) + 1)
                
                # Ordenar
                propuesta_final['DateStatus_Order'] = propuesta_final['DateStatus'].map(
                    {'PAST_DUE': 1, 'CURRENT+FUTURE': 2}
                )
                propuesta_final = propuesta_final.sort_values(by=['DateStatus_Order', 'ItemNo']).drop(columns='DateStatus_Order')
                
                # Guardar archivo Excel
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_file = os.path.join(self.desktop, f"GRUPOS_KITEO_{timestamp}.xlsx")
                
                self.update_status("💾 Guardando archivo Excel...")
                
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    propuesta_cols = ['Group_ID', 'ItemNo', 'Description', 'Entity', 'WO_Count',
                                    'WO_String', 'OpenQty_Total', 'QtyFcst_Total', 'DateStatus',
                                    'ReqDate_Min', 'ReqDate_Max', 'Rev', 'UM']
                    propuesta_final[propuesta_cols].to_excel(writer, sheet_name='Propuesta_Kiteo', index=False)
                
                # Actualizar datos en la interfaz
                self.current_data = propuesta_final
                self.update_metrics(df, propuesta_final)
                self.update_data_table(propuesta_final)
                
                self.update_status(f"🎉 PROCESO COMPLETADO - Archivo: GRUPOS_KITEO_{timestamp}.xlsx")
                self.show_snackbar(f"✅ Archivo guardado exitosamente", ft.Colors.GREEN_600)
                
                # Intentar abrir el archivo
                try:
                    os.startfile(output_file)
                except:
                    self.show_snackbar("⚠️ Abrir manualmente el archivo desde el escritorio", ft.Colors.ORANGE_600)
                
            except Exception as e:
                self.update_status(f"❌ ERROR: {str(e)}")
                self.show_snackbar(f"❌ Error: {str(e)}", ft.Colors.RED_600)
            finally:
                self.is_processing = False
        
        # Ejecutar en hilo separado
        processing_thread = threading.Thread(target=run_processing, daemon=True)
        processing_thread.start()
    
    def update_status(self, message):
        """Actualizar el estado en la interfaz"""
        if hasattr(self, 'status_text'):
            self.status_text.value = message
            self.page.update()
        print(message)  # También imprimir en consola
    
    def update_metrics(self, df_original, df_propuesta):
        """Actualizar las métricas en las tarjetas"""
        if not hasattr(self, 'metric_cards'):
            return
        
        # Calcular métricas
        total_records = len(df_original)
        total_groups = len(df_propuesta)
        past_due = len(df_propuesta[df_propuesta['DateStatus'] == 'PAST_DUE'])
        current_future = len(df_propuesta[df_propuesta['DateStatus'] == 'CURRENT+FUTURE'])
        total_qty = df_propuesta['OpenQty_Total'].sum()
        
        # Actualizar tarjetas
        metrics = [
            {"value": f"{total_records:,}", "label": "Registros Totales"},
            {"value": f"{total_groups:,}", "label": "Grupos Generados"},
            {"value": f"{past_due:,}", "label": "Past Due"},
            {"value": f"{current_future:,}", "label": "Current+Future"},
        ]
        
        for i, metric in enumerate(metrics):
            if i < len(self.metric_cards):
                # Encontrar el Text widget del valor
                card_content = self.metric_cards[i].content
                if hasattr(card_content, 'controls'):
                    for control in card_content.controls:
                        if isinstance(control, ft.Text) and control.size == 32:
                            control.value = metric["value"]
                            break
        
        self.page.update()
    
    def update_data_table(self, df):
        """Actualizar la tabla de datos"""
        if not hasattr(self, 'data_table'):
            return
        
        # Limpiar filas existentes
        self.data_table.rows.clear()
        
        # Agregar las primeras 10 filas como ejemplo
        for i, row in df.head(10).iterrows():
            self.data_table.rows.append(
                ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(row['Group_ID']))),
                    ft.DataCell(ft.Text(str(row['ItemNo']))),
                    ft.DataCell(ft.Text(str(row['Entity']))),
                    ft.DataCell(ft.Text(str(row['WO_Count']))),
                    ft.DataCell(ft.Text(str(row['OpenQty_Total']))),
                    ft.DataCell(ft.Text(str(row['DateStatus']))),
                ])
            )
        
        self.page.update()
    
    def show_dashboard(self):
        """Mostrar el dashboard principal"""
        self.page.clean()
        
        # Header
        header = ft.Container(
            content=ft.Row([
                ft.Column([
                    ft.Text("🚀 SISTEMA DE PROPUESTA DE KITEO", 
                           size=32, weight=ft.FontWeight.BOLD, color=self.theme_color),
                    ft.Text(f"Generación automática de grupos - {datetime.now().strftime('%B %d, %Y')}", 
                           size=16, color=ft.Colors.WHITE54),
                ]),
                ft.Container(expand=True),
                ft.ElevatedButton(
                    "▶️ EJECUTAR PROCESO",
                    icon=ft.Icons.PLAY_ARROW,
                    bgcolor=self.theme_color,
                    color=ft.Colors.WHITE,
                    height=50,
                    width=200,
                    on_click=self.process_kiteo_data,
                    style=ft.ButtonStyle(
                        shape=ft.RoundedRectangleBorder(radius=10)
                    )
                )
            ]),
            padding=ft.padding.only(bottom=30),
        )
        
        # Crear tarjetas de métricas
        self.metric_cards = [
            self.create_metric_card("Registros Totales", "0", ft.Icons.STORAGE, ft.Colors.BLUE),
            self.create_metric_card("Grupos Generados", "0", ft.Icons.GROUP_WORK, ft.Colors.GREEN),
            self.create_metric_card("Past Due", "0", ft.Icons.WARNING, ft.Colors.RED),
            self.create_metric_card("Current+Future", "0", ft.Icons.SCHEDULE, ft.Colors.ORANGE),
        ]
        
        # Estado del proceso
        self.status_text = ft.Text(
            "🔄 Listo para procesar datos...",
            size=16,
            color=ft.Colors.WHITE70,
        )
        
        status_container = ft.Container(
            content=ft.Column([
                ft.Text("📊 Estado del Proceso", size=20, weight=ft.FontWeight.W_500),
                ft.Container(height=10),
                self.status_text,
            ]),
            padding=30,
            bgcolor=ft.Colors.GREY_800,
            border_radius=15,
        )
        
        # Información de la base de datos
        db_info = ft.Container(
            content=ft.Column([
                ft.Text("🗃️ Información de Base de Datos", size=20, weight=ft.FontWeight.W_500),
                ft.Container(height=10),
                ft.Text(f"📁 Ruta: {self.db_path}", size=12, color=ft.Colors.WHITE70),
                ft.Text(f"💾 Salida: {self.desktop}", size=12, color=ft.Colors.WHITE70),
                ft.Text("🔍 Consulta: Registros con OpenQty > 0 y WO válida", size=12, color=ft.Colors.WHITE70),
                ft.Text("📅 Agrupación: PAST_DUE separado, CURRENT+FUTURE juntos", size=12, color=ft.Colors.WHITE70),
            ]),
            padding=30,
            bgcolor=ft.Colors.BLUE_GREY_800,
            border_radius=15,
        )
        
        # Tabla de vista previa
        self.data_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("ID")),
                ft.DataColumn(ft.Text("ItemNo")),
                ft.DataColumn(ft.Text("Entity")),
                ft.DataColumn(ft.Text("WO Count")),
                ft.DataColumn(ft.Text("Open Qty")),
                ft.DataColumn(ft.Text("Status")),
            ],
            rows=[],
        )
        
        preview_container = ft.Container(
            content=ft.Column([
                ft.Text("📋 Vista Previa de Resultados", size=20, weight=ft.FontWeight.W_500),
                ft.Container(height=10),
                ft.Text("(Se mostrará después del procesamiento)", size=12, color=ft.Colors.WHITE54),
                ft.Container(height=15),
                ft.Container(
                    content=self.data_table,
                    bgcolor=ft.Colors.GREY_900,
                    border_radius=10,
                    padding=10,
                ),
            ]),
            padding=30,
            bgcolor=ft.Colors.GREY_800,
            border_radius=15,
            expand=True,
        )
        
        # Layout principal
        main_content = ft.Column([
            header,
            # Métricas
            ft.Row(self.metric_cards, spacing=20),
            ft.Container(height=20),
            # Estado y DB Info
            ft.Row([
                ft.Container(content=status_container, expand=True),
                ft.Container(width=20),
                ft.Container(content=db_info, expand=True),
            ]),
            ft.Container(height=20),
            # Vista previa
            preview_container,
        ], expand=True)
        
        self.page.add(main_content)
    
    def create_metric_card(self, title, value, icon, color):
        """Crear tarjeta de métrica"""
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(icon, color=color, size=30),
                    ft.Container(expand=True),
                ]),
                ft.Container(height=20),
                ft.Text(value, size=32, weight=ft.FontWeight.BOLD),
                ft.Text(title, size=16, color=ft.Colors.WHITE54),
            ]),
            bgcolor=ft.Colors.GREY_800,
            padding=25,
            border_radius=15,
            expand=True,
            animate=ft.Animation(300, ft.AnimationCurve.EASE_IN_OUT),
        )


# Función principal para ejecutar el dashboard
def main():
    app = KiteoDashboard()
    ft.app(target=app.main)

if __name__ == "__main__":
    main()