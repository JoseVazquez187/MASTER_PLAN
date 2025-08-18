import flet as ft
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import os
from typing import Dict, List, Tuple
import json

class PartValidationApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Part Number Validation Dashboard"
        self.page.theme_mode = "light"
        self.page.window_width = 1400
        self.page.window_height = 900
        
        # Database path
        self.db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        
        # Data storage
        self.loaded_data = None
        self.validation_results = []
        
        # UI Components
        self.file_picker = ft.FilePicker(on_result=self.on_file_picked)
        self.page.overlay.append(self.file_picker)
        
        # Status indicators
        self.status_text = ft.Text("Listo para cargar archivo", color="blue")
        self.progress_bar = ft.ProgressBar(width=400, visible=False)
        
        # File upload section
        self.upload_button = ft.ElevatedButton(
            "📁 Cargar Archivo (CSV/Excel)",
            icon=ft.Icons.UPLOAD_FILE,
            on_click=lambda _: self.file_picker.pick_files(
                dialog_title="Seleccionar archivo",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["csv", "xlsx", "xls"]
            )
        )
        
        # Validate button
        self.validate_button = ft.ElevatedButton(
            "🔍 Validar Part Numbers",
            icon=ft.Icons.SEARCH,
            disabled=True,
            on_click=self.validate_parts
        )
        
        # Export button
        self.export_button = ft.ElevatedButton(
            "📊 Exportar Resultados",
            icon=ft.Icons.DOWNLOAD,
            disabled=True,
            on_click=self.export_results
        )
        
        # Filter dropdown
        self.filter_dropdown = ft.Dropdown(
            label="Filtrar por estado",
            width=200,
            options=[
                ft.dropdown.Option("all", "Todos"),
                ft.dropdown.Option("valid", "✅ Válidos"),
                ft.dropdown.Option("red_flag", "⚠️ Red Flag"),
                ft.dropdown.Option("no_bom", "🚫 Sin BOM"),
                ft.dropdown.Option("no_history", "📋 Sin Historial"),
            ],
            value="all",
            on_change=self.filter_results
        )
        
        # Results data table
        self.results_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Estado")),
                ft.DataColumn(ft.Text("Item No")),
                ft.DataColumn(ft.Text("Req Qty")),
                ft.DataColumn(ft.Text("Req Date")),
                ft.DataColumn(ft.Text("BOM Status")),
                ft.DataColumn(ft.Text("Invoice History")),
                ft.DataColumn(ft.Text("Observaciones")),
                ft.DataColumn(ft.Text("Acciones")),
            ],
            rows=[],
        )
        
        # Summary cards
        self.summary_cards = ft.Row([])
        
        self.build_ui()
    
    def build_ui(self):
        """Construye la interfaz de usuario"""
        
        # Header
        header = ft.Container(
            content=ft.Column([
                ft.Text(
                    "Dashboard de Validación de Part Numbers",
                    size=24,
                    weight=ft.FontWeight.BOLD,
                    color="blue"
                ),
                ft.Text(
                    "Valida números de parte contra historial de facturación y BOM",
                    size=14,
                    color="grey"
                ),
            ]),
            padding=20,
            bgcolor="#f5f5f5",
            border_radius=10,
            margin=ft.margin.only(bottom=20)
        )
        
        # Upload section
        upload_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("1. Cargar Archivo", size=18, weight=ft.FontWeight.BOLD),
                    ft.Text("Formato requerido: ItemNo, ReqQty, ReqDate", color="grey"),
                    ft.Row([
                        self.upload_button,
                        self.validate_button,
                        self.export_button,
                    ]),
                    self.status_text,
                    self.progress_bar,
                ]),
                padding=20,
            )
        )
        
        # Filter section
        filter_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("2. Filtros y Resumen", size=18, weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.filter_dropdown,
                        ft.Container(width=20),
                        self.summary_cards,
                    ]),
                ]),
                padding=20,
            )
        )
        
        # Results section
        results_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("3. Resultados de Validación", size=18, weight=ft.FontWeight.BOLD),
                    ft.Container(
                        content=ft.Column([
                            self.results_table,
                        ], scroll=ft.ScrollMode.AUTO),
                        height=400,
                    )
                ]),
                padding=20,
            )
        )
        
        # Main layout
        self.page.add(
            ft.Column([
                header,
                upload_section,
                filter_section,
                results_section,
            ], scroll=ft.ScrollMode.AUTO)
        )
    
    def on_file_picked(self, e: ft.FilePickerResultEvent):
        """Maneja la selección de archivo"""
        if e.files:
            file_path = e.files[0].path
            self.load_file(file_path)
    
    def load_file(self, file_path: str):
        """Carga el archivo seleccionado"""
        try:
            self.status_text.value = "Cargando archivo..."
            self.status_text.color = "orange"
            self.page.update()
            
            # Determinar tipo de archivo y cargar
            if file_path.endswith('.csv'):
                self.loaded_data = pd.read_csv(file_path)
            else:
                self.loaded_data = pd.read_excel(file_path)
            
            # Validar columnas requeridas
            required_columns = ['ItemNo', 'ReqQty', 'ReqDate']
            missing_columns = [col for col in required_columns if col not in self.loaded_data.columns]
            
            if missing_columns:
                raise ValueError(f"Columnas faltantes: {', '.join(missing_columns)}")
            
            # Limpiar y preparar datos
            self.loaded_data['ReqDate'] = pd.to_datetime(self.loaded_data['ReqDate'])
            self.loaded_data = self.loaded_data.dropna(subset=['ItemNo'])
            
            self.status_text.value = f"✅ Archivo cargado exitosamente. {len(self.loaded_data)} registros encontrados."
            self.status_text.color = "green"
            self.validate_button.disabled = False
            
        except Exception as ex:
            self.status_text.value = f"❌ Error al cargar archivo: {str(ex)}"
            self.status_text.color = "red"
            self.validate_button.disabled = True
        
        self.page.update()
    
    def validate_parts(self, e):
        """Ejecuta la validación de part numbers"""
        if self.loaded_data is None:
            return
        
        self.progress_bar.visible = True
        self.status_text.value = "Validando part numbers..."
        self.status_text.color = "orange"
        self.page.update()
        
        try:
            self.validation_results = []
            total_parts = len(self.loaded_data)
            
            for index, row in self.loaded_data.iterrows():
                # Actualizar progress bar
                progress = (index + 1) / total_parts
                self.progress_bar.value = progress
                self.page.update()
                
                item_no = row['ItemNo']
                req_qty = row['ReqQty']
                req_date = row['ReqDate']
                
                # Consultar BOM
                bom_data = self.get_bom_data(item_no)
                
                # Consultar historial de invoices
                invoice_history = self.get_invoice_history(item_no)
                
                # Determinar status de validación
                validation_result = self.determine_validation_status(
                    item_no, req_qty, req_date, bom_data, invoice_history
                )
                
                self.validation_results.append(validation_result)
            
            self.update_results_display()
            self.update_summary_cards()
            
            self.status_text.value = f"✅ Validación completada. {len(self.validation_results)} items procesados."
            self.status_text.color = "green"
            self.export_button.disabled = False
            
        except Exception as ex:
            self.status_text.value = f"❌ Error en validación: {str(ex)}"
            self.status_text.color = "red"
        finally:
            self.progress_bar.visible = False
            self.page.update()
    
    def get_bom_data(self, item_no: str) -> Dict:
        """Obtiene datos de BOM para un item"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                # Buscar como key (producto final)
                query_key = """
                SELECT Component, Unit_Qty, Level_Number, Description
                FROM bom 
                WHERE key = ?
                """
                key_results = pd.read_sql_query(query_key, conn, params=[item_no])
                
                # Buscar como component (componente)
                query_component = """
                SELECT key, Unit_Qty, Level_Number, Description
                FROM bom 
                WHERE Component = ?
                """
                component_results = pd.read_sql_query(query_component, conn, params=[item_no])
                
                return {
                    'as_key': key_results.to_dict('records') if not key_results.empty else [],
                    'as_component': component_results.to_dict('records') if not component_results.empty else [],
                    'has_bom': not key_results.empty or not component_results.empty
                }
        except Exception:
            return {'as_key': [], 'as_component': [], 'has_bom': False}
    
    def get_invoice_history(self, item_no: str) -> Dict:
        """Obtiene historial de invoices para un item"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                query = """
                SELECT Item_Number, S_Qty, Inv_Dt, Price, Customer, Ord_Dt
                FROM invoiced 
                WHERE Item_Number = ?
                ORDER BY Inv_Dt DESC
                LIMIT 50
                """
                results = pd.read_sql_query(query, conn, params=[item_no])
                
                if not results.empty:
                    # Calcular estadísticas
                    total_invoiced = results['S_Qty'].sum()
                    avg_qty = results['S_Qty'].mean()
                    last_invoice = results.iloc[0]['Inv_Dt']
                    unique_customers = results['Customer'].nunique()
                    
                    return {
                        'has_history': True,
                        'records': results.to_dict('records'),
                        'total_invoiced': total_invoiced,
                        'avg_qty': avg_qty,
                        'last_invoice': last_invoice,
                        'unique_customers': unique_customers,
                        'invoice_count': len(results)
                    }
                else:
                    return {
                        'has_history': False,
                        'records': [],
                        'total_invoiced': 0,
                        'avg_qty': 0,
                        'last_invoice': None,
                        'unique_customers': 0,
                        'invoice_count': 0
                    }
        except Exception:
            return {
                'has_history': False,
                'records': [],
                'total_invoiced': 0,
                'avg_qty': 0,
                'last_invoice': None,
                'unique_customers': 0,
                'invoice_count': 0
            }
    
    def determine_validation_status(self, item_no: str, req_qty: float, req_date: datetime, 
                                  bom_data: Dict, invoice_history: Dict) -> Dict:
        """Determina el estado de validación para un item"""
        
        red_flags = []
        observations = []
        
        # Verificar BOM
        bom_status = "Sin BOM"
        if bom_data['has_bom']:
            if bom_data['as_key']:
                bom_status = f"Producto final ({len(bom_data['as_key'])} componentes)"
            elif bom_data['as_component']:
                bom_status = f"Componente (usado en {len(bom_data['as_component'])} productos)"
        else:
            red_flags.append("Sin estructura BOM")
        
        # Verificar historial
        history_status = "Sin historial"
        if invoice_history['has_history']:
            history_status = f"{invoice_history['invoice_count']} facturas"
            
            # Verificar si la cantidad solicitada es inusual
            if req_qty > invoice_history['avg_qty'] * 3:
                red_flags.append("Cantidad solicitada muy alta vs historial")
            
            # Verificar última facturación
            if invoice_history['last_invoice']:
                last_date = pd.to_datetime(invoice_history['last_invoice'])
                days_since_last = (datetime.now() - last_date).days
                if days_since_last > 365:
                    red_flags.append("No facturado en más de 1 año")
        else:
            red_flags.append("Sin historial de facturación")
        
        # Determinar estado general
        if red_flags:
            status = "⚠️ Red Flag"
            status_color = "red"
        elif not bom_data['has_bom']:
            status = "🚫 Sin BOM"
            status_color = "orange"
        elif not invoice_history['has_history']:
            status = "📋 Sin Historial"
            status_color = "orange"
        else:
            status = "✅ Válido"
            status_color = "green"
        
        # Compilar observaciones
        if red_flags:
            observations.extend(red_flags)
        if invoice_history['has_history']:
            observations.append(f"Promedio facturado: {invoice_history['avg_qty']:.1f}")
        
        return {
            'item_no': item_no,
            'req_qty': req_qty,
            'req_date': req_date.strftime('%Y-%m-%d'),
            'status': status,
            'status_color': status_color,
            'bom_status': bom_status,
            'history_status': history_status,
            'observations': '; '.join(observations) if observations else 'N/A',
            'bom_data': bom_data,
            'invoice_history': invoice_history,
            'red_flags': red_flags
        }
    
    def update_results_display(self):
        """Actualiza la tabla de resultados"""
        self.results_table.rows.clear()
        
        filtered_results = self.get_filtered_results()
        
        for result in filtered_results:
            # Botón de detalles
            details_button = ft.IconButton(
                icon=ft.Icons.INFO,
                tooltip="Ver detalles",
                on_click=lambda e, r=result: self.show_details(r)
            )
            
            row = ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(result['status'])),
                    ft.DataCell(ft.Text(str(result['item_no']))),
                    ft.DataCell(ft.Text(str(result['req_qty']))),
                    ft.DataCell(ft.Text(result['req_date'])),
                    ft.DataCell(ft.Text(result['bom_status'])),
                    ft.DataCell(ft.Text(result['history_status'])),
                    ft.DataCell(ft.Text(
                        result['observations'][:50] + "..." if len(result['observations']) > 50 
                        else result['observations']
                    )),
                    ft.DataCell(details_button),
                ]
            )
            self.results_table.rows.append(row)
        
        self.page.update()
    
    def get_filtered_results(self) -> List[Dict]:
        """Obtiene resultados filtrados según selección"""
        if not self.validation_results:
            return []
        
        filter_value = self.filter_dropdown.value
        
        if filter_value == "all":
            return self.validation_results
        elif filter_value == "valid":
            return [r for r in self.validation_results if r['status'] == "✅ Válido"]
        elif filter_value == "red_flag":
            return [r for r in self.validation_results if r['status'] == "⚠️ Red Flag"]
        elif filter_value == "no_bom":
            return [r for r in self.validation_results if r['status'] == "🚫 Sin BOM"]
        elif filter_value == "no_history":
            return [r for r in self.validation_results if r['status'] == "📋 Sin Historial"]
        
        return self.validation_results
    
    def filter_results(self, e):
        """Maneja el cambio de filtro"""
        self.update_results_display()
    
    def update_summary_cards(self):
        """Actualiza las tarjetas de resumen"""
        if not self.validation_results:
            return
        
        total = len(self.validation_results)
        valid = len([r for r in self.validation_results if r['status'] == "✅ Válido"])
        red_flag = len([r for r in self.validation_results if r['status'] == "⚠️ Red Flag"])
        no_bom = len([r for r in self.validation_results if r['status'] == "🚫 Sin BOM"])
        no_history = len([r for r in self.validation_results if r['status'] == "📋 Sin Historial"])
        
        cards = [
            self.create_summary_card("Total", str(total), "blue"),
            self.create_summary_card("Válidos", str(valid), "green"),
            self.create_summary_card("Red Flags", str(red_flag), "red"),
            self.create_summary_card("Sin BOM", str(no_bom), "orange"),
            self.create_summary_card("Sin Historial", str(no_history), "grey"),
        ]
        
        self.summary_cards.controls = cards
        self.page.update()
    
    def create_summary_card(self, title: str, value: str, color: str) -> ft.Card:
        """Crea una tarjeta de resumen"""
        return ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text(value, size=24, weight=ft.FontWeight.BOLD, color=color),
                    ft.Text(title, size=12),
                ], alignment=ft.MainAxisAlignment.CENTER),
                width=100,
                height=80,
                padding=10,
            )
        )
    
    def show_details(self, result: Dict):
        """Muestra detalles de un item en un dialog"""
        
        # Preparar contenido de detalles
        bom_content = []
        if result['bom_data']['as_key']:
            bom_content.append(ft.Text("Como producto final:", weight=ft.FontWeight.BOLD))
            for comp in result['bom_data']['as_key'][:5]:  # Mostrar solo primeros 5
                bom_content.append(ft.Text(f"  • {comp['Component']} (Qty: {comp['Unit_Qty']})"))
        
        if result['bom_data']['as_component']:
            bom_content.append(ft.Text("Usado en productos:", weight=ft.FontWeight.BOLD))
            for prod in result['bom_data']['as_component'][:5]:  # Mostrar solo primeros 5
                bom_content.append(ft.Text(f"  • {prod['key']} (Qty: {prod['Unit_Qty']})"))
        
        history_content = []
        if result['invoice_history']['has_history']:
            hist = result['invoice_history']
            history_content.extend([
                ft.Text(f"Total facturado: {hist['total_invoiced']}"),
                ft.Text(f"Cantidad promedio: {hist['avg_qty']:.1f}"),
                ft.Text(f"Última factura: {hist['last_invoice']}"),
                ft.Text(f"Clientes únicos: {hist['unique_customers']}"),
            ])
        
        dialog = ft.AlertDialog(
            title=ft.Text(f"Detalles: {result['item_no']}"),
            content=ft.Container(
                content=ft.Column([
                    ft.Text("📋 Información BOM:", weight=ft.FontWeight.BOLD),
                    *bom_content,
                    ft.Divider(),
                    ft.Text("📊 Historial de Facturas:", weight=ft.FontWeight.BOLD),
                    *history_content,
                    ft.Divider(),
                    ft.Text("⚠️ Observaciones:", weight=ft.FontWeight.BOLD),
                    ft.Text(result['observations']),
                ], scroll=ft.ScrollMode.AUTO),
                width=500,
                height=400,
            ),
            actions=[
                ft.TextButton("Cerrar", on_click=lambda e: self.close_dialog())
            ]
        )
        
        self.page.dialog = dialog
        dialog.open = True
        self.page.update()
    
    def close_dialog(self):
        """Cierra el dialog actual"""
        self.page.dialog.open = False
        self.page.update()
    
    def export_results(self, e):
        """Exporta los resultados a Excel"""
        if not self.validation_results:
            return
        
        try:
            # Preparar datos para exportar
            export_data = []
            for result in self.validation_results:
                export_data.append({
                    'Item_No': result['item_no'],
                    'Req_Qty': result['req_qty'],
                    'Req_Date': result['req_date'],
                    'Status': result['status'],
                    'BOM_Status': result['bom_status'],
                    'History_Status': result['history_status'],
                    'Observations': result['observations'],
                    'Red_Flags': len(result['red_flags']),
                })
            
            df = pd.DataFrame(export_data)
            
            # Generar nombre de archivo
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"validation_results_{timestamp}.xlsx"
            
            # Exportar
            df.to_excel(filename, index=False)
            
            self.status_text.value = f"✅ Resultados exportados a: {filename}"
            self.status_text.color = "green"
            
        except Exception as ex:
            self.status_text.value = f"❌ Error al exportar: {str(ex)}"
            self.status_text.color = "red"
        
        self.page.update()

def main(page: ft.Page):
    app = PartValidationApp(page)

if __name__ == "__main__":
    ft.app(target=main)