#!/usr/bin/env python3
"""
R4 Database to Google Sheets Sync - Aplicación Completa
Sincroniza tablas SQLite de R4Database con Google Sheets usando Flet
"""

import flet as ft
import sqlite3
import pandas as pd
import os
import json
from datetime import datetime, timedelta
import asyncio
from pathlib import Path
import logging
import shutil
import webbrowser
import platform
# Configuración principal
DEFAULT_DB_PATH = r"C:\Users\J.Vazquez\Desktop\R4Database.db"  # Tu ruta específica
DEFAULT_CSV_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "R4_CSV_Export")
LOOKER_STUDIO_URL = "https://lookerstudio.google.com/navigation/reporting"

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DatabaseValidator:
    """Valida y obtiene información de la base de datos"""
    
    def __init__(self, db_path):
        self.db_path = db_path
        
    def validate_connection(self):
        """Valida la conexión a la base de datos"""
        try:
            if not os.path.exists(self.db_path):
                return False, f"Base de datos no encontrada: {self.db_path}"
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Test de conectividad
            cursor.execute("SELECT 1")
            cursor.fetchone()
            
            conn.close()
            return True, "Conexión exitosa"
            
        except Exception as e:
            return False, f"Error de conexión: {str(e)}"
    
    def get_all_tables(self):
        """Obtiene todas las tablas disponibles en la base de datos"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT name FROM sqlite_master 
                WHERE type='table' AND name NOT LIKE 'sqlite_%'
                ORDER BY name
            """)
            
            tables = [row[0] for row in cursor.fetchall()]
            conn.close()
            
            return True, tables
            
        except Exception as e:
            return False, f"Error obteniendo tablas: {str(e)}"
    
    def get_table_info(self, table_name):
        """Obtiene información detallada de una tabla"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Escapar nombre de tabla con comillas si es necesario
            safe_table_name = f'"{table_name}"' if ' ' in table_name or '-' in table_name else table_name
            
            # Obtener número de registros
            try:
                cursor.execute(f"SELECT COUNT(*) FROM {safe_table_name}")
                record_count = cursor.fetchone()[0]
            except sqlite3.Error:
                # Intentar con backticks
                cursor.execute(f"SELECT COUNT(*) FROM `{table_name}`")
                record_count = cursor.fetchone()[0]
            
            # Obtener información de columnas
            try:
                cursor.execute(f"PRAGMA table_info({safe_table_name})")
                columns_info = cursor.fetchall()
            except sqlite3.Error:
                cursor.execute(f"PRAGMA table_info(`{table_name}`)")
                columns_info = cursor.fetchall()
            
            # Obtener muestra de datos
            try:
                cursor.execute(f"SELECT * FROM {safe_table_name} LIMIT 3")
                sample_data = cursor.fetchall()
            except sqlite3.Error:
                cursor.execute(f"SELECT * FROM `{table_name}` LIMIT 3")
                sample_data = cursor.fetchall()
            
            # Obtener tamaño estimado (en MB)
            try:
                cursor.execute(f"SELECT SUM(LENGTH(CAST(* AS TEXT))) FROM {safe_table_name} LIMIT 1000")
                size_sample = cursor.fetchone()[0] or 0
            except sqlite3.Error:
                size_sample = 0
            
            estimated_size_mb = (size_sample * record_count / 1000) / (1024 * 1024) if record_count > 0 else 0
            
            conn.close()
            
            return True, {
                "record_count": record_count,
                "column_count": len(columns_info),
                "columns": [col[1] for col in columns_info],  # Nombres de columnas
                "column_types": [(col[1], col[2]) for col in columns_info],  # (nombre, tipo)
                "sample_data": sample_data,
                "estimated_size_mb": round(estimated_size_mb, 2)
            }
            
        except Exception as e:
            # Log del error específico para debugging
            logger.warning(f"Error obteniendo info de tabla '{table_name}': {str(e)}")
            return False, f"Error: {str(e)[:50]}..."
    
    def export_table_to_csv(self, table_name, output_path, limit=None):
        """Exporta una tabla a archivo CSV"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            # Construir query
            query = f"SELECT * FROM `{table_name}`"
            if limit and limit > 0:
                query += f" LIMIT {limit}"
            
            # Leer datos con pandas
            df = pd.read_sql_query(query, conn)
            conn.close()
            
            if df.empty:
                return False, "La tabla está vacía", 0
            
            # Crear directorio si no existe
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Exportar a CSV
            df.to_csv(output_path, index=False, encoding='utf-8')
            
            file_size_mb = os.path.getsize(output_path) / (1024 * 1024)
            
            return True, f"Exportado exitosamente: {file_size_mb:.2f} MB", len(df)
            
        except Exception as e:
            return False, f"Error exportando {table_name}: {str(e)}", 0
        
class CSVManager:
    """Gestiona la creación y organización de archivos CSV"""
    
    def __init__(self, csv_folder):
        self.csv_folder = Path(csv_folder)
        
    def setup_folder_structure(self):
        """Crea la estructura de carpetas necesaria"""
        try:
            # Solo carpeta principal
            self.csv_folder.mkdir(parents=True, exist_ok=True)
            
            # Carpeta de respaldos (opcional)
            (self.csv_folder / "backups").mkdir(exist_ok=True)
            
            # Crear archivo README
            readme_content = f"""# R4 CSV Export Folder
            
Carpeta creada automáticamente el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Archivos CSV:
Todos los archivos se guardan directamente en esta carpeta:
- expedite.csv
- sales_order_table.csv  
- oh.csv
- openorder.csv
- actionmessages.csv
- vendor.csv
- bom.csv
- woinquiry.csv
- (y otras tablas según selección)

## Subcarpetas:
- **backups/**: Respaldos automáticos de archivos anteriores

## Para Looker Studio:
1. Sube estos archivos CSV a Google Drive
2. Conecta Looker Studio a Google Drive
3. Selecciona los archivos CSV como fuente de datos
4. Los archivos se actualizan automáticamente

Última actualización: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
            
            (self.csv_folder / "README.txt").write_text(readme_content)
            
            return True, f"Carpeta creada en: {self.csv_folder}"
            
        except Exception as e:
            return False, f"Error creando carpeta: {str(e)}"
    
    def get_category_for_table(self, table_name):
        """Determina la categoría de una tabla"""
        table_categories = {
            "inventory": ["oh", "in521", "in92", "whs_location", "qa_oh", "reworkloc"],
            "production": ["sales_order_table", "woinquiry", "bom", "operation_wo", "openwo"],
            "purchasing": ["expedite", "openorder", "vendor", "po351", "transit"],
            "planning": ["fcst", "actionmessages", "safety_stock", "kpt", "pr561"],
            "quality": ["qa_oh", "reworkloc", "kiting_groups"],
            "finance": ["credit_memos", "prepetual_in56"],
            "master_data": ["entity", "holiday", "table_yearweek", "where_use"]
        }
        
        table_lower = table_name.lower()
        
        for category, tables in table_categories.items():
            if any(keyword in table_lower for keyword in tables):
                return category
        
        return "master_data"  # Por defecto
    
    def get_csv_path(self, table_name):
        """Obtiene la ruta completa para el archivo CSV de una tabla"""
        # Todos los archivos CSV van directamente en la carpeta principal
        filename = f"{table_name}.csv"
        return self.csv_folder / filename
    
    def backup_existing_file(self, csv_path):
        """Crea respaldo de archivo existente"""
        try:
            if csv_path.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_name = f"{csv_path.stem}_backup_{timestamp}.csv"
                backup_path = self.csv_folder / "backups" / backup_name
                
                shutil.copy2(csv_path, backup_path)
                return True, str(backup_path)
            
            return True, "No necesita backup"
            
        except Exception as e:
            return False, f"Error creando backup: {str(e)}"
        
class LookerStudioHelper:
    """Ayuda con la integración a Looker Studio"""
    
    @staticmethod
    def create_integration_guide(csv_folder):
        """Crea una guía de integración con Looker Studio"""
        guide_content = f"""# 📊 Guía de Integración con Looker Studio

## 🚀 Pasos para Conectar:

### 1. Preparar Archivos
✅ Archivos CSV generados en: {csv_folder}
✅ Todos los archivos en una carpeta principal (sin subcarpetas)
✅ Archivos listos para Looker Studio

### 2. Subir a Google Drive (Recomendado)
1. Abre Google Drive: https://drive.google.com
2. Crea una carpeta llamada "R4_Data"
3. Sube todos los archivos CSV de {os.path.basename(csv_folder)}
4. Comparte la carpeta (opcional, para colaboración)

### 3. Conectar Looker Studio
1. Ve a Looker Studio: https://lookerstudio.google.com
2. Clic en "Crear" → "Fuente de datos"
3. Selecciona "Google Drive"
4. Navega a tu carpeta R4_Data
5. Selecciona el archivo CSV deseado

### 4. Configurar Fuente de Datos
- **Tipo de archivo**: CSV
- **Codificación**: UTF-8
- **Separador**: Coma (,)
- **Primera fila**: Encabezados ✅

### 5. Opciones Avanzadas
- **Actualización automática**: Configurar refresh cada hora/día
- **Esquema**: Revisar tipos de datos detectados
- **Filtros**: Aplicar si es necesario

## 📋 Archivos CSV Principales:

### 🎯 Archivos Prioritarios:
- **expedite.csv** - Lista de expedite (crítico para compras)
- **sales_order_table.csv** - Órdenes de venta (ventas y producción)
- **oh.csv** - Inventario actual (stock disponible)
- **openorder.csv** - Órdenes abiertas (compras pendientes)
- **actionmessages.csv** - Mensajes de acción (planeación)

### 📊 Archivos Adicionales:
- **vendor.csv** - Proveedores
- **bom.csv** - Lista de materiales
- **woinquiry.csv** - Órdenes de trabajo
- **in521.csv** - Inventario detallado
- **fcst.csv** - Pronósticos
- (y otros según tu selección)

## 🔄 Actualización de Datos

### Automática (Recomendado):
1. Programa la aplicación para ejecutarse diariamente
2. Los archivos CSV se actualizan automáticamente en la misma ubicación
3. Looker Studio detecta los cambios automáticamente

### Manual:
1. Ejecuta la aplicación cuando necesites
2. Los archivos se sobrescriben con datos actuales
3. Refresca los reportes en Looker Studio

## 📁 Ubicación de Archivos:
- **Carpeta principal**: {csv_folder}
- **Backups**: {csv_folder}/backups/
- **Todos los CSV**: Directamente en la carpeta principal (fácil acceso)

## 💡 Tips para Looker Studio:

1. **Tipos de Datos**:
   - Fechas: Asegúrate que se detecten como DATE
   - Números: Configura como NUMBER o CURRENCY
   - Texto: Usa TEXT para descripciones

2. **Performance**:
   - Usa filtros a nivel de fuente de datos
   - Considera agregaciones pre-calculadas para tablas grandes

3. **Visualizaciones Recomendadas**:
   - **expedite.csv**: Tabla con items críticos, gráficos de barras por proveedor
   - **sales_order_table.csv**: Líneas de tiempo, KPIs de ventas
   - **oh.csv**: Scorecards de inventario, tablas de stock bajo

## 🔗 Enlaces Útiles:
- Looker Studio: https://lookerstudio.google.com
- Google Drive: https://drive.google.com
- Documentación: https://developers.google.com/looker-studio

---
Generado automáticamente el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Estructura simplificada: Todos los archivos CSV en una sola carpeta
"""
        
        guide_path = Path(csv_folder) / "LOOKER_STUDIO_GUIDE.md"
        guide_path.write_text(guide_content, encoding='utf-8')
        
        return str(guide_path)
    
class R4CSVExporterApp:
    
    def __init__(self, page: ft.Page):
        self.page = page
        self.db_validator = None
        self.csv_manager = CSVManager(DEFAULT_CSV_FOLDER)
        self.available_tables = []
        self.selected_tables = set()
        
        # UI Components - Tab 1
        self.db_path_field = None
        self.csv_folder_field = None
        self.tables_list = None
        self.status_text = None
        self.progress_bar = None
        self.db_stats_text = None
        self.export_button = None
        
        # UI Components - Tab 2 (Queries)
        self.query_text = None
        self.query_name_field = None
        self.saved_queries_dropdown = None
        self.query_status_text = None
        self.query_progress_bar = None
        
        # Query management
        self.queries_file = Path(DEFAULT_CSV_FOLDER) / "saved_queries.json"
        self.saved_queries = self.load_saved_queries()
        
        self.setup_page()
        self.build_ui()
    
    def setup_page(self):
        """Configura la página principal"""
        self.page.title = "R4 Database → CSV Exporter para Looker Studio"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.window_width = 1100
        self.page.window_height = 800
        self.page.window_resizable = True
        self.page.padding = 20
    
    def build_ui(self):
        """Construye la interfaz de usuario con pestañas"""
        
        # Header
        header = ft.Container(
            content=ft.Column([
                ft.Text(
                    "📊 R4 Database → CSV Exporter",
                    size=28,
                    weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_700
                ),
                ft.Text(
                    "Exporta tablas de R4Database a CSV para Looker Studio",
                    size=16,
                    color=ft.Colors.GREY_600
                ),
                ft.Text(
                    "✨ Validación automática • Selección de tablas • Queries personalizadas",
                    size=12,
                    color=ft.Colors.BLUE_600
                )
            ]),
            margin=ft.margin.only(bottom=20)
        )
        
        # Crear pestañas
        tabs = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=[
                ft.Tab(
                    text="📊 Exportar Tablas",
                    icon=ft.Icons.TABLE_VIEW,
                    content=self.build_tables_tab()
                ),
                ft.Tab(
                    text="⚡ Queries Personalizadas",
                    icon=ft.Icons.CODE,
                    content=self.build_queries_tab()
                )
            ],
            expand=1
        )
        
        # Layout principal con scroll
        main_content = ft.Column(
            [header, tabs],
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )
        
        self.page.add(main_content)
        
        # Auto-validate if default path exists
        if os.path.exists(DEFAULT_DB_PATH):
            self.validate_database(None)

    def build_tables_tab(self):
        """Construye la pestaña de exportación de tablas"""
        
        # Database Configuration
        self.db_path_field = ft.TextField(
            label="📁 Ruta de la base de datos R4",
            value=DEFAULT_DB_PATH,
            expand=True,
            hint_text="Selecciona el archivo R4Database.db"
        )
        
        self.db_stats_text = ft.Text("Base de datos no validada", color=ft.Colors.GREY_600)
        
        db_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("🔗 Configuración de Base de Datos", size=16, weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.db_path_field,
                        ft.ElevatedButton(
                            "Explorar",
                            icon=ft.Icons.FOLDER_OPEN,
                            on_click=self.pick_database_file
                        ),
                        ft.ElevatedButton(
                            "Validar BD",
                            icon=ft.Icons.VERIFIED,
                            on_click=self.validate_database,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.GREEN_600,
                                color=ft.Colors.WHITE
                            )
                        )
                    ]),
                    self.db_stats_text
                ]),
                padding=15
            )
        )
        
        # CSV Output Configuration
        self.csv_folder_field = ft.TextField(
            label="📂 Carpeta de salida CSV",
            value=DEFAULT_CSV_FOLDER,
            expand=True,
            hint_text="Carpeta donde se guardarán los archivos CSV"
        )
        
        csv_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("📁 Configuración de Salida CSV", size=16, weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.csv_folder_field,
                        ft.ElevatedButton(
                            "Explorar",
                            icon=ft.Icons.FOLDER_OPEN,
                            on_click=self.pick_csv_folder
                        ),
                        ft.ElevatedButton(
                            "Crear Estructura",
                            icon=ft.Icons.CREATE_NEW_FOLDER,
                            on_click=self.create_folder_structure,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.PURPLE_600,
                                color=ft.Colors.WHITE
                            )
                        )
                    ]),
                    ft.Text(
                        "💡 Todos los archivos CSV se guardan directamente en la carpeta principal",
                        size=12,
                        color=ft.Colors.GREY_600
                    )
                ]),
                padding=15
            )
        )
        
        # Tables Selection
        self.tables_list = ft.Column(
            scroll=ft.ScrollMode.AUTO, 
            height=280,
            spacing=5,
            tight=True
        )
        
        tables_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Text("📋 Selección de Tablas", size=16, weight=ft.FontWeight.BOLD),
                        ft.ElevatedButton(
                            "Actualizar Lista",
                            icon=ft.Icons.REFRESH,
                            on_click=self.refresh_tables_list,
                            style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_600, color=ft.Colors.WHITE)
                        )
                    ]),
                    ft.Row([
                        ft.TextButton("✅ Seleccionar Todas", on_click=self.select_all_tables),
                        ft.TextButton("❌ Deseleccionar Todas", on_click=self.deselect_all_tables),
                        ft.TextButton("⭐ Solo Principales", on_click=self.select_main_tables),
                        ft.TextButton("📊 Solo con Datos", on_click=self.select_tables_with_data)
                    ]),
                    self.tables_list
                ]),
                padding=15
            )
        )
        
        # Export Section
        self.progress_bar = ft.ProgressBar(visible=False)
        self.status_text = ft.Text("Listo para exportar", color=ft.Colors.GREY_600)
        
        self.export_button = ft.ElevatedButton(
            "🚀 Exportar Tablas Seleccionadas",
            icon=ft.Icons.DOWNLOAD,
            on_click=self.start_export,
            style=ft.ButtonStyle(
                bgcolor=ft.Colors.ORANGE_600,
                color=ft.Colors.WHITE
            ),
            disabled=True
        )
        
        export_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("📤 Exportación", size=16, weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.export_button,
                        ft.ElevatedButton(
                            "📊 Abrir Looker Studio",
                            icon=ft.Icons.OPEN_IN_NEW,
                            on_click=self.open_looker_studio,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.DEEP_PURPLE_600,
                                color=ft.Colors.WHITE
                            )
                        ),
                        ft.ElevatedButton(
                            "📁 Abrir Carpeta CSV",
                            icon=ft.Icons.FOLDER,
                            on_click=self.open_csv_folder,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.BROWN_600,
                                color=ft.Colors.WHITE
                            )
                        )
                    ]),
                    self.progress_bar,
                    self.status_text
                ]),
                padding=15
            )
        )
        
        return ft.Column([
            db_section,
            csv_section,
            tables_section,
            export_section
        ], scroll=ft.ScrollMode.AUTO, expand=True)

    def build_queries_tab(self):
        """Construye la pestaña de queries personalizadas"""
        
        # Query Editor Section
        self.query_text = ft.TextField(
            label="📝 Query SQL",
            multiline=True,
            max_lines=8,
            hint_text="Escribe tu query SQL aquí...\nEjemplo: SELECT * FROM expedite WHERE Vendor = 'ACME' LIMIT 100",
            value="-- Ejemplo: Items en expedite por proveedor\nSELECT ItemNo, Description, Vendor, ReqDate, ReqQty \nFROM expedite \nWHERE Vendor LIKE '%ACME%' \nORDER BY ReqDate DESC \nLIMIT 50;",
            expand=True
        )
        
        # Query Name and Save
        self.query_name_field = ft.TextField(
            label="📋 Nombre del Query",
            hint_text="Ej: Items_Expedite_ACME",
            expand=True
        )
        
        query_editor_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("⚡ Editor de Queries SQL", size=16, weight=ft.FontWeight.BOLD),
                    self.query_text,
                    ft.Row([
                        self.query_name_field,
                        ft.ElevatedButton(
                            "💾 Guardar Query",
                            icon=ft.Icons.SAVE,
                            on_click=self.save_query,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.GREEN_600,
                                color=ft.Colors.WHITE
                            )
                        )
                    ])
                ]),
                padding=15
            )
        )
        
        # Saved Queries Section
        self.saved_queries_dropdown = ft.Dropdown(
            label="📚 Queries Guardadas",
            hint_text="Selecciona un query guardado...",
            expand=True,
            on_change=self.load_selected_query
        )
        
        self.refresh_queries_dropdown()
        
        saved_queries_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("📚 Queries Guardadas", size=16, weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.saved_queries_dropdown,
                        ft.ElevatedButton(
                            "🔄 Cargar",
                            icon=ft.Icons.UPLOAD,
                            on_click=self.load_selected_query,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.BLUE_600,
                                color=ft.Colors.WHITE
                            )
                        ),
                        ft.ElevatedButton(
                            "🗑️ Eliminar",
                            icon=ft.Icons.DELETE,
                            on_click=self.delete_selected_query,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.RED_600,
                                color=ft.Colors.WHITE
                            )
                        )
                    ])
                ]),
                padding=15
            )
        )
        
        # Query Execution Section
        self.query_progress_bar = ft.ProgressBar(visible=False)
        self.query_status_text = ft.Text("Listo para ejecutar query", color=ft.Colors.GREY_600)
        
        execution_section = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("🚀 Ejecución y Exportación", size=16, weight=ft.FontWeight.BOLD),
                    ft.Row([
                        ft.ElevatedButton(
                            "▶️ Ejecutar Query",
                            icon=ft.Icons.PLAY_ARROW,
                            on_click=self.execute_query,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.ORANGE_600,
                                color=ft.Colors.WHITE
                            )
                        ),
                        ft.ElevatedButton(
                            "📥 Ejecutar y Exportar CSV",
                            icon=ft.Icons.DOWNLOAD,
                            on_click=self.execute_and_export_query,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.DEEP_ORANGE_600,
                                color=ft.Colors.WHITE
                            )
                        ),
                        ft.ElevatedButton(
                            "📁 Abrir Carpeta",
                            icon=ft.Icons.FOLDER,
                            on_click=self.open_csv_folder,
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.BROWN_600,
                                color=ft.Colors.WHITE
                            )
                        )
                    ]),
                    self.query_progress_bar,
                    self.query_status_text,
                    ft.Text(
                        "💡 Tips: Usa LIMIT para queries grandes • Los CSV se guardan en la carpeta principal",
                        size=12,
                        color=ft.Colors.GREY_600
                    )
                ]),
                padding=15
            )
        )
        
        # Query Examples
        examples_section = ft.Card(
            content=ft.Container(
                content=ft.ExpansionTile(
                    title=ft.Text("📖 Ejemplos de Queries", weight=ft.FontWeight.BOLD),
                    controls=[
                        ft.Column([
                            ft.Text("🔍 Items en Expedite por Proveedor:", weight=ft.FontWeight.BOLD, size=12),
                            ft.Container(
                                content=ft.Text(
                                    "SELECT ItemNo, Description, Vendor, ReqDate\nFROM expedite\nWHERE Vendor LIKE '%ACME%'\nORDER BY ReqDate;",
                                    size=11,
                                    color=ft.Colors.GREY_700
                                ),
                                bgcolor=ft.Colors.GREY_100,
                                padding=10,
                                border_radius=5
                            ),
                            
                            ft.Text("📦 Inventario Bajo por Item:", weight=ft.FontWeight.BOLD, size=12),
                            ft.Container(
                                content=ft.Text(
                                    "SELECT ItemNo, Description, OH\nFROM oh\nWHERE CAST(OH AS INTEGER) < 100\nORDER BY CAST(OH AS INTEGER);",
                                    size=11,
                                    color=ft.Colors.GREY_700
                                ),
                                bgcolor=ft.Colors.GREY_100,
                                padding=10,
                                border_radius=5
                            ),
                            
                            ft.Text("🏭 Órdenes de Venta por Cliente:", weight=ft.FontWeight.BOLD, size=12),
                            ft.Container(
                                content=ft.Text(
                                    "SELECT Cust_Name, COUNT(*) as Total_Orders, SUM(CAST(OpenValue AS REAL)) as Total_Value\nFROM sales_order_table\nGROUP BY Cust_Name\nORDER BY Total_Value DESC;",
                                    size=11,
                                    color=ft.Colors.GREY_700
                                ),
                                bgcolor=ft.Colors.GREY_100,
                                padding=10,
                                border_radius=5
                            )
                        ], spacing=10)
                    ]
                ),
                padding=15
            )
        )
        
        return ft.Column([
            query_editor_section,
            saved_queries_section,
            execution_section,
            examples_section
        ], scroll=ft.ScrollMode.AUTO, expand=True)

    def validate_database(self, e):
        """Valida la conexión a la base de datos"""
        db_path = self.db_path_field.value.strip()
        
        if not db_path:
            self.show_error("Por favor selecciona una base de datos")
            return
        
        self.status_text.value = "🔍 Validando base de datos..."
        self.page.update()
        
        self.db_validator = DatabaseValidator(db_path)
        success, message = self.db_validator.validate_connection()
        
        if success:
            # Obtener lista de tablas
            tables_success, tables_result = self.db_validator.get_all_tables()
            
            if tables_success:
                self.available_tables = tables_result
                self.load_tables_list()
                
                # Estadísticas
                file_size = os.path.getsize(db_path) / (1024 * 1024)
                self.db_stats_text.value = f"✅ Conectado: {len(tables_result)} tablas, {file_size:.2f} MB"
                self.db_stats_text.color = ft.Colors.GREEN_600
                
                self.status_text.value = f"✅ Base de datos validada: {len(tables_result)} tablas encontradas"
                self.status_text.color = ft.Colors.GREEN_600
                
                self.export_button.disabled = False
            else:
                self.show_error(f"Error obteniendo tablas: {tables_result}")
        else:
            self.show_error(f"Error de validación: {message}")
            self.db_stats_text.value = "❌ Error de conexión"
            self.db_stats_text.color = ft.Colors.RED_600
        
        self.page.update()
    
    async def pick_database_file(self, e):
        """Selecciona archivo de base de datos"""
        def handle_file_picker_result(e: ft.FilePickerResultEvent):
            if e.files:
                self.db_path_field.value = e.files[0].path
                self.page.update()
        
        file_picker = ft.FilePicker(on_result=handle_file_picker_result)
        self.page.overlay.append(file_picker)
        self.page.update()
        
        await file_picker.pick_files(
            dialog_title="Seleccionar base de datos R4",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["db", "sqlite", "sqlite3"]
        )
    
    async def pick_csv_folder(self, e):
        """Selecciona carpeta de salida CSV"""
        def handle_folder_picker_result(e: ft.FilePickerResultEvent):
            if e.path:
                self.csv_folder_field.value = e.path
                self.csv_manager = CSVManager(e.path)
                self.page.update()
        
        folder_picker = ft.FilePicker(on_result=handle_folder_picker_result)
        self.page.overlay.append(folder_picker)
        self.page.update()
        
        await folder_picker.get_directory_path(dialog_title="Seleccionar carpeta para archivos CSV")

    def create_folder_structure(self, e):
        """Crea la estructura de carpetas CSV"""
        csv_folder = self.csv_folder_field.value.strip()
        
        if not csv_folder:
            self.show_error("Por favor especifica una carpeta de salida")
            return
        
        self.csv_manager = CSVManager(csv_folder)
        success, message = self.csv_manager.setup_folder_structure()
        
        if success:
            self.status_text.value = "✅ Estructura de carpetas creada"
            self.status_text.color = ft.Colors.GREEN_600
            
            # Crear guía de Looker Studio
            guide_path = LookerStudioHelper.create_integration_guide(csv_folder)
            
            self.show_info(
                "Estructura Creada",
                f"✅ Carpeta CSV creada exitosamente\n\n"
                f"📁 Ubicación: {csv_folder}\n"
                f"📄 Archivos CSV: Se guardan directamente en la carpeta principal\n"
                f"💾 Backups: Subcarpeta 'backups' para respaldos\n"
                f"📖 Guía: README.txt y LOOKER_STUDIO_GUIDE.md\n\n"
                f"🔗 Todos los archivos estarán en un solo lugar, fáciles de encontrar"
            )
        else:
            self.show_error(f"Error creando estructura: {message}")
        
        self.page.update()
    
    def load_tables_list(self):
        """Carga la lista de tablas disponibles"""
        if not self.available_tables:
            return
        
        self.tables_list.controls.clear()
        
        # Obtener información de cada tabla
        for table_name in self.available_tables:
            success, info = self.db_validator.get_table_info(table_name)
            
            if success:
                record_count = info['record_count']
                column_count = info['column_count']
                size_mb = info['estimated_size_mb']
                
                # Determinar categoría
                category = self.csv_manager.get_category_for_table(table_name)
                category_Icons = {
                    "inventory": "📦",
                    "production": "🏭", 
                    "purchasing": "🛒",
                    "planning": "📋",
                    "quality": "🔍",
                    "finance": "💰",
                    "master_data": "🗂️"
                }
                category_icon = category_Icons.get(category, "📄")
                
                subtitle = f"{category_icon} {record_count:,} registros • {column_count} columnas • ~{size_mb:.1f} MB"
                
                # Determinar si es tabla principal
                main_tables = ["sales_order_table", "expedite", "openorder", "oh", "actionmessages", "vendor", "bom"]
                is_main = table_name.lower() in [t.lower() for t in main_tables]
                title_color = ft.Colors.BLUE_700 if is_main else None
                
                if is_main:
                    subtitle += " ⭐"
                    
            else:
                # Error obteniendo información - pero podemos intentar exportar
                subtitle = f"⚠️ {info[:60]}..."
                title_color = ft.Colors.ORANGE_600
                subtitle += "\n(Se puede intentar exportar)"
            
            checkbox = ft.Checkbox(
                label=table_name,
                value=False,
                on_change=self.on_table_selected
            )
            
            table_tile = ft.Container(
                content=ft.Row([
                    checkbox,
                    ft.Column([
                        ft.Text(table_name, weight=ft.FontWeight.BOLD, color=title_color, size=14),
                        ft.Text(subtitle, size=11, color=ft.Colors.GREY_400)  # ← También cambiar este color
                    ], spacing=2, expand=True)
                ], spacing=10),
                padding=ft.padding.symmetric(vertical=5, horizontal=10),
                margin=ft.margin.symmetric(vertical=2),
                border_radius=5,
                bgcolor=ft.Colors.GREY_800,  # ← FONDO OSCURO PARA TEMA OSCURO
                border=ft.border.all(1, ft.Colors.GREY_600),  # ← BORDE SUTIL
                on_click=lambda e, table=table_name: self.toggle_table_selection(table)
            )
            
            self.tables_list.controls.append(table_tile)
        
        self.page.update()
    
    def refresh_tables_list(self, e):
        """Actualiza la lista de tablas"""
        if self.db_validator:
            self.validate_database(None)
        else:
            self.show_warning("Primero valida la base de datos")
    
    def toggle_table_selection(self, table_name):
        """Alterna la selección de una tabla"""
        for control in self.tables_list.controls:
            if isinstance(control, ft.Container):
                # Buscar el checkbox dentro del Container
                row = control.content
                if isinstance(row, ft.Row) and len(row.controls) > 0:
                    checkbox = row.controls[0]
                    if isinstance(checkbox, ft.Checkbox) and checkbox.label == table_name:
                        checkbox.value = not checkbox.value
                        self.on_table_selected(None)
                        break
        self.page.update()
    
    def on_table_selected(self, e):
        """Maneja la selección de tablas"""
        self.selected_tables.clear()
        
        for control in self.tables_list.controls:
            if isinstance(control, ft.Container):
                # Buscar el checkbox dentro del Container
                row = control.content
                if isinstance(row, ft.Row) and len(row.controls) > 0:
                    checkbox = row.controls[0]
                    if isinstance(checkbox, ft.Checkbox) and checkbox.value:
                        self.selected_tables.add(checkbox.label)
        
        count = len(self.selected_tables)
        if count > 0:
            self.status_text.value = f"📋 {count} tabla(s) seleccionada(s) para exportar"
            self.status_text.color = ft.Colors.BLUE_600
        else:
            self.status_text.value = "Selecciona tablas para exportar"
            self.status_text.color = ft.Colors.GREY_600
        
        self.page.update()
    
    def select_all_tables(self, e):
        """Selecciona todas las tablas"""
        for control in self.tables_list.controls:
            if isinstance(control, ft.Container):
                row = control.content
                if isinstance(row, ft.Row) and len(row.controls) > 0:
                    checkbox = row.controls[0]
                    if isinstance(checkbox, ft.Checkbox):
                        checkbox.value = True
        
        self.on_table_selected(None)
    
    def deselect_all_tables(self, e):
        """Deselecciona todas las tablas"""
        for control in self.tables_list.controls:
            if isinstance(control, ft.Container):
                row = control.content
                if isinstance(row, ft.Row) and len(row.controls) > 0:
                    checkbox = row.controls[0]
                    if isinstance(checkbox, ft.Checkbox):
                        checkbox.value = False
        
        self.on_table_selected(None)
    
    def select_main_tables(self, e):
        """Selecciona solo las tablas principales"""
        main_tables = ["sales_order_table", "expedite", "openorder", "oh", "actionmessages", "vendor", "bom"]
        
        for control in self.tables_list.controls:
            if isinstance(control, ft.Container):
                row = control.content
                if isinstance(row, ft.Row) and len(row.controls) > 0:
                    checkbox = row.controls[0]
                    if isinstance(checkbox, ft.Checkbox):
                        checkbox.value = checkbox.label.lower() in [t.lower() for t in main_tables]
        
        self.on_table_selected(None)
    
    def select_tables_with_data(self, e):
        """Selecciona solo tablas que tienen datos"""
        for control in self.tables_list.controls:
            if isinstance(control, ft.Container):
                row = control.content
                if isinstance(row, ft.Row) and len(row.controls) > 0:
                    checkbox = row.controls[0]
                    if isinstance(checkbox, ft.Checkbox):
                        table_name = checkbox.label
                        
                        success, info = self.db_validator.get_table_info(table_name)
                        checkbox.value = success and info.get('record_count', 0) > 0
        
        self.on_table_selected(None)

    # ================================
    # QUERY MANAGEMENT FUNCTIONS
    # ================================
    
    def load_saved_queries(self):
        """Carga las queries guardadas desde archivo JSON"""
        try:
            if self.queries_file.exists():
                with open(self.queries_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            logger.error(f"Error cargando queries: {e}")
            return {}
    
    def save_queries_to_file(self):
        """Guarda las queries al archivo JSON"""
        try:
            # Crear directorio si no existe
            self.queries_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.queries_file, 'w', encoding='utf-8') as f:
                json.dump(self.saved_queries, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            logger.error(f"Error guardando queries: {e}")
            return False
    
    def refresh_queries_dropdown(self):
        """Actualiza el dropdown con las queries guardadas"""
        try:
            if hasattr(self, 'saved_queries_dropdown') and self.saved_queries_dropdown:
                # Limpiar opciones existentes
                self.saved_queries_dropdown.options.clear()
                
                # Agregar nuevas opciones
                for query_name in self.saved_queries.keys():
                    self.saved_queries_dropdown.options.append(
                        ft.dropdown.Option(key=query_name, text=query_name)
                    )
                
                # Forzar actualización del dropdown
                self.saved_queries_dropdown.value = None  # Limpiar selección
                
                # Actualizar la página
                if hasattr(self, 'page'):
                    self.page.update()
                    
        except Exception as e:
            logger.error(f"Error refrescando dropdown: {e}")

    def save_query(self, e):
        """Guarda un query con nombre"""
        query_name = self.query_name_field.value.strip()
        query_text = self.query_text.value.strip()
        
        if not query_name:
            self.show_query_error("Por favor ingresa un nombre para el query")
            return
        
        if not query_text:
            self.show_query_error("Por favor ingresa el texto del query")
            return
        
        # Guardar query
        self.saved_queries[query_name] = {
            "query": query_text,
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "description": f"Query personalizado: {query_name}"
        }
        
        if self.save_queries_to_file():
            self.refresh_queries_dropdown()
            self.query_status_text.value = f"✅ Query '{query_name}' guardado exitosamente"
            self.query_status_text.color = ft.Colors.GREEN_600
            
            # Limpiar campos
            self.query_name_field.value = ""
            
        else:
            self.show_query_error("Error guardando el query")
        
        self.page.update()
    
    def load_selected_query(self, e):
        """Carga el query seleccionado en el editor"""
        if not self.saved_queries_dropdown.value:
            return
        
        query_name = self.saved_queries_dropdown.value
        if query_name in self.saved_queries:
            query_data = self.saved_queries[query_name]
            self.query_text.value = query_data["query"]
            self.query_name_field.value = query_name
            
            self.query_status_text.value = f"📝 Query '{query_name}' cargado"
            self.query_status_text.color = ft.Colors.BLUE_600
            
            self.page.update()
    
    def delete_selected_query(self, e):
        """Elimina el query seleccionado"""
        if not self.saved_queries_dropdown.value:
            self.query_status_text.value = "❌ Selecciona un query para eliminar"
            self.query_status_text.color = ft.Colors.RED_600
            self.page.update()
            return
        
        query_name = self.saved_queries_dropdown.value
        
        try:
            if query_name in self.saved_queries:
                # Eliminar directamente sin confirmación
                del self.saved_queries[query_name]
                
                if self.save_queries_to_file():
                    # Limpiar selección ANTES de refrescar
                    self.saved_queries_dropdown.value = None
                    
                    # Refrescar dropdown
                    self.refresh_queries_dropdown()
                    
                    # Limpiar también el editor si tenía ese query cargado
                    if self.query_name_field.value == query_name:
                        self.query_text.value = ""
                        self.query_name_field.value = ""
                    
                    # Mensaje de éxito
                    self.query_status_text.value = f"🗑️ Query '{query_name}' eliminado exitosamente"
                    self.query_status_text.color = ft.Colors.GREEN_600
                        
                else:
                    self.query_status_text.value = "❌ Error guardando cambios al archivo"
                    self.query_status_text.color = ft.Colors.RED_600
            else:
                self.query_status_text.value = "❌ Query no encontrado"
                self.query_status_text.color = ft.Colors.RED_600
                
        except Exception as ex:
            self.query_status_text.value = f"❌ Error eliminando query: {str(ex)}"
            self.query_status_text.color = ft.Colors.RED_600
            logger.error(f"Error eliminando query: {ex}")
        
        # Forzar actualización completa
        self.page.update()

    def execute_query(self, e):
        """Ejecuta el query y muestra resultados"""
        query_text = self.query_text.value.strip()
        
        if not query_text:
            self.show_query_error("Por favor ingresa un query para ejecutar")
            return
        
        if not self.db_validator:
            self.show_query_error("Primero valida la base de datos en la pestaña principal")
            return
        
        self.query_progress_bar.visible = True
        self.query_status_text.value = "⏳ Ejecutando query..."
        self.query_status_text.color = ft.Colors.BLUE_600
        self.page.update()
        
        try:
            conn = sqlite3.connect(self.db_validator.db_path)
            
            # Ejecutar query
            df = pd.read_sql_query(query_text, conn)
            conn.close()
            
            if df.empty:
                self.query_status_text.value = "⚠️ Query ejecutado pero no devolvió resultados"
                self.query_status_text.color = ft.Colors.ORANGE_600
            else:
                records = len(df)
                columns = len(df.columns)
                self.query_status_text.value = f"✅ Query ejecutado: {records:,} registros, {columns} columnas"
                self.query_status_text.color = ft.Colors.GREEN_600
                
                # Mostrar muestra de resultados
                self.show_query_results(df.head(10))
            
        except Exception as e:
            self.query_status_text.value = f"❌ Error en query: {str(e)[:100]}..."
            self.query_status_text.color = ft.Colors.RED_600
            logger.error(f"Error ejecutando query: {e}")
        
        finally:
            self.query_progress_bar.visible = False
            self.page.update()
    
    async def execute_and_export_query(self, e):
        """Ejecuta el query y exporta directamente a CSV"""
        query_text = self.query_text.value.strip()
        
        if not query_text:
            self.show_query_error("Por favor ingresa un query para ejecutar")
            return
        
        if not self.db_validator:
            self.show_query_error("Primero valida la base de datos en la pestaña principal")
            return
        
        # Generar nombre de archivo
        query_name = self.query_name_field.value.strip()
        if not query_name:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"custom_query_{timestamp}.csv"
        else:
            # Limpiar nombre para uso como archivo
            safe_name = "".join(c for c in query_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filename = f"{safe_name}.csv"
        
        self.query_progress_bar.visible = True
        self.query_status_text.value = "⏳ Ejecutando y exportando query..."
        self.query_status_text.color = ft.Colors.BLUE_600
        self.page.update()
        
        try:
            # Crear carpeta CSV si no existe
            csv_folder = self.csv_folder_field.value.strip() or DEFAULT_CSV_FOLDER
            self.csv_manager = CSVManager(csv_folder)
            success, message = self.csv_manager.setup_folder_structure()
            
            if not success:
                self.show_query_error(f"Error configurando carpeta: {message}")
                return
            
            # Ejecutar query
            conn = sqlite3.connect(self.db_validator.db_path)
            df = pd.read_sql_query(query_text, conn)
            conn.close()
            
            if df.empty:
                self.query_status_text.value = "⚠️ Query no devolvió resultados para exportar"
                self.query_status_text.color = ft.Colors.ORANGE_600
                return
            
            # Exportar a CSV
            csv_path = Path(csv_folder) / filename
            
            # Crear backup si existe
            if csv_path.exists():
                backup_success, backup_result = self.csv_manager.backup_existing_file(csv_path)
                if backup_success and "backup_" in str(backup_result):
                    logger.info(f"Backup creado: {backup_result}")
            
            # Guardar CSV
            df.to_csv(csv_path, index=False, encoding='utf-8')
            
            file_size = csv_path.stat().st_size / (1024 * 1024)  # MB
            records = len(df)
            
            self.query_status_text.value = f"✅ Exportado: {records:,} registros → {filename} ({file_size:.2f} MB)"
            self.query_status_text.color = ft.Colors.GREEN_600
            
            # Mostrar diálogo de éxito
            self.show_info(
                "Query Exportado",
                f"🎉 Query ejecutado y exportado exitosamente\n\n"
                f"📁 Archivo: {filename}\n"
                f"📊 Registros: {records:,}\n"
                f"💾 Tamaño: {file_size:.2f} MB\n"
                f"📂 Ubicación: {csv_folder}\n\n"
                f"🔗 El archivo está listo para Looker Studio"
            )
            
        except Exception as e:
            self.query_status_text.value = f"❌ Error: {str(e)[:100]}..."
            self.query_status_text.color = ft.Colors.RED_600
            logger.error(f"Error ejecutando y exportando query: {e}")
        
        finally:
            self.query_progress_bar.visible = False
            self.page.update()
    
    def show_query_results(self, df_sample):
        """Muestra una muestra de los resultados del query"""
        try:
            # Crear tabla para mostrar resultados
            columns = list(df_sample.columns)
            rows = []
            
            for _, row in df_sample.iterrows():
                row_data = [str(row[col])[:50] + "..." if len(str(row[col])) > 50 else str(row[col]) for col in columns]
                rows.append(row_data)
            
            # Crear contenido del diálogo
            table_content = f"Columnas: {', '.join(columns[:5])}{'...' if len(columns) > 5 else ''}\n\n"
            table_content += f"Muestra de {len(rows)} filas:\n"
            
            for i, row in enumerate(rows[:5]):
                table_content += f"Fila {i+1}: {', '.join(row[:3])}{'...' if len(row) > 3 else ''}\n"
            
            if len(rows) > 5:
                table_content += "..."
            
            dialog = ft.AlertDialog(
                title=ft.Text("📊 Resultados del Query"),
                content=ft.Container(
                    content=ft.Text(table_content, size=12),
                    width=500,
                    height=300
                ),
                actions=[ft.TextButton("OK", on_click=lambda e: self.close_dialog())]
            )
            
            self.page.dialog = dialog
            dialog.open = True
            self.page.update()
            
        except Exception as e:
            logger.error(f"Error mostrando resultados: {e}")
    
    def show_query_error(self, message):
        """Muestra un mensaje de error específico para queries"""
        self.query_status_text.value = f"❌ {message}"
        self.query_status_text.color = ft.Colors.RED_600
        self.page.update()
        
        dialog = ft.AlertDialog(
            title=ft.Text("❌ Error en Query"),
            content=ft.Text(message),
            actions=[ft.TextButton("OK", on_click=lambda e: self.close_dialog())]
        )
        
        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

    # ================================
    # EXPORT FUNCTIONS (Continuación desde queries)
    # ================================
    
    async def start_export(self, e):
        """Inicia el proceso de exportación de tablas seleccionadas"""
        if not self.selected_tables:
            self.show_error("Por favor selecciona al menos una tabla para exportar")
            return
        
        if not self.db_validator:
            self.show_error("Primero valida la base de datos")
            return
        
        csv_folder = self.csv_folder_field.value.strip()
        if not csv_folder:
            self.show_error("Por favor especifica una carpeta de salida")
            return
        
        # Configurar UI para exportación
        self.export_button.disabled = True
        self.progress_bar.visible = True
        self.progress_bar.value = 0
        self.page.update()
        
        try:
            # Crear estructura de carpetas
            self.csv_manager = CSVManager(csv_folder)
            success, message = self.csv_manager.setup_folder_structure()
            
            if not success:
                self.show_error(f"Error configurando carpeta: {message}")
                return
            
            # Exportar cada tabla seleccionada
            total_tables = len(self.selected_tables)
            exported_count = 0
            failed_tables = []
            export_summary = []
            
            for i, table_name in enumerate(self.selected_tables):
                # Actualizar progreso
                progress = (i / total_tables)
                self.progress_bar.value = progress
                self.status_text.value = f"📤 Exportando {table_name}... ({i+1}/{total_tables})"
                self.status_text.color = ft.Colors.BLUE_600
                self.page.update()
                
                # Obtener ruta CSV
                csv_path = self.csv_manager.get_csv_path(table_name)
                
                # Crear backup si existe
                backup_success, backup_result = self.csv_manager.backup_existing_file(csv_path)
                if backup_success and "backup_" in str(backup_result):
                    logger.info(f"Backup creado para {table_name}: {backup_result}")
                
                # Exportar tabla
                export_success, export_message, record_count = self.db_validator.export_table_to_csv(
                    table_name, str(csv_path)
                )
                
                if export_success:
                    exported_count += 1
                    file_size = csv_path.stat().st_size / (1024 * 1024) if csv_path.exists() else 0
                    export_summary.append({
                        "table": table_name,
                        "records": record_count,
                        "size_mb": file_size,
                        "status": "✅ Éxito"
                    })
                    logger.info(f"Exportado exitosamente: {table_name} ({record_count:,} registros)")
                else:
                    failed_tables.append(f"{table_name}: {export_message}")
                    export_summary.append({
                        "table": table_name,
                        "records": 0,
                        "size_mb": 0,
                        "status": f"❌ Error: {export_message[:30]}..."
                    })
                    logger.error(f"Error exportando {table_name}: {export_message}")
                
                # Pequeña pausa para no sobrecargar
                await asyncio.sleep(0.1)
            
            # Progreso completo
            self.progress_bar.value = 1.0
            
            # Crear guía de Looker Studio
            guide_path = LookerStudioHelper.create_integration_guide(csv_folder)
            
            # Mostrar resumen
            if exported_count == total_tables:
                self.status_text.value = f"🎉 Exportación completa: {exported_count} tablas exportadas"
                self.status_text.color = ft.Colors.GREEN_600
                
                self.show_export_summary("Exportación Exitosa", export_summary, csv_folder)
            else:
                self.status_text.value = f"⚠️ Exportación parcial: {exported_count}/{total_tables} tablas"
                self.status_text.color = ft.Colors.ORANGE_600
                
                self.show_export_summary("Exportación Parcial", export_summary, csv_folder, failed_tables)
                
        except Exception as e:
            self.status_text.value = f"❌ Error durante exportación: {str(e)[:50]}..."
            self.status_text.color = ft.Colors.RED_600
            logger.error(f"Error durante exportación: {e}")
            self.show_error(f"Error durante exportación: {str(e)}")
        
        finally:
            # Restaurar UI
            self.export_button.disabled = False
            self.progress_bar.visible = False
            self.page.update()
    
    def show_export_summary(self, title, export_summary, csv_folder, failed_tables=None):
        """Muestra un resumen detallado de la exportación"""
        
        # Crear contenido del resumen
        summary_text = f"📊 Resumen de Exportación\n\n"
        summary_text += f"📁 Carpeta: {csv_folder}\n\n"
        
        total_records = 0
        total_size = 0
        success_count = 0
        
        for item in export_summary:
            status_icon = "✅" if "Éxito" in item["status"] else "❌"
            summary_text += f"{status_icon} {item['table']}: {item['records']:,} registros, {item['size_mb']:.2f} MB\n"
            
            if "Éxito" in item["status"]:
                total_records += item['records']
                total_size += item['size_mb']
                success_count += 1
        
        summary_text += f"\n📈 Totales:\n"
        summary_text += f"• Tablas exitosas: {success_count}/{len(export_summary)}\n"
        summary_text += f"• Registros exportados: {total_records:,}\n"
        summary_text += f"• Tamaño total: {total_size:.2f} MB\n"
        
        if failed_tables:
            summary_text += f"\n❌ Tablas con errores:\n"
            for error in failed_tables[:5]:  # Mostrar máximo 5 errores
                summary_text += f"• {error}\n"
            if len(failed_tables) > 5:
                summary_text += f"• ... y {len(failed_tables) - 5} más\n"
        
        summary_text += f"\n🔗 Próximos pasos:\n"
        summary_text += f"1. Los archivos CSV están listos para Looker Studio\n"
        summary_text += f"2. Revisa la guía LOOKER_STUDIO_GUIDE.md\n"
        summary_text += f"3. Sube los archivos a Google Drive\n"
        summary_text += f"4. Conecta Looker Studio a tus datos\n"
        
        # Crear diálogo
        dialog = ft.AlertDialog(
            title=ft.Text(f"🎉 {title}"),
            content=ft.Container(
                content=ft.Column([
                    ft.Text(summary_text, size=12),
                    ft.Row([
                        ft.ElevatedButton(
                            "📁 Abrir Carpeta",
                            icon=ft.Icons.FOLDER,
                            on_click=lambda e: [self.open_csv_folder(e), self.close_dialog()],
                            style=ft.ButtonStyle(bgcolor=ft.Colors.BROWN_600, color=ft.Colors.WHITE)
                        ),
                        ft.ElevatedButton(
                            "📊 Looker Studio",
                            icon=ft.Icons.OPEN_IN_NEW,
                            on_click=lambda e: [self.open_looker_studio(e), self.close_dialog()],
                            style=ft.ButtonStyle(bgcolor=ft.Colors.DEEP_PURPLE_600, color=ft.Colors.WHITE)
                        )
                    ])
                ], scroll=ft.ScrollMode.AUTO),
                width=600,
                height=400
            ),
            actions=[
                ft.TextButton("Cerrar", on_click=lambda e: self.close_dialog())
            ]
        )
        
        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

    # ================================
    # UI HELPER FUNCTIONS
    # ================================
    
    def open_looker_studio(self, e):
        """Abre Looker Studio en el navegador"""
        try:
            webbrowser.open(LOOKER_STUDIO_URL)
            self.status_text.value = "🌐 Looker Studio abierto en el navegador"
            self.status_text.color = ft.Colors.BLUE_600
            self.page.update()
        except Exception as ex:
            self.show_error(f"Error abriendo Looker Studio: {str(ex)}")
    
    def open_csv_folder(self, e):
        """Abre la carpeta CSV en el explorador de archivos"""
        try:
            csv_folder = self.csv_folder_field.value.strip() or DEFAULT_CSV_FOLDER
            
            if not os.path.exists(csv_folder):
                self.show_warning(f"La carpeta no existe: {csv_folder}")
                return
            
            # Abrir carpeta según el sistema operativo
            import platform
            system = platform.system()
            
            if system == "Windows":
                os.startfile(csv_folder)
            elif system == "Darwin":  # macOS
                os.system(f"open '{csv_folder}'")
            else:  # Linux
                os.system(f"xdg-open '{csv_folder}'")
            
            self.status_text.value = f"📁 Carpeta CSV abierta: {os.path.basename(csv_folder)}"
            self.status_text.color = ft.Colors.BLUE_600
            self.page.update()
            
        except Exception as ex:
            self.show_error(f"Error abriendo carpeta: {str(ex)}")
    
    def close_dialog(self):
        """Cierra cualquier diálogo abierto"""
        try:
            if hasattr(self.page, 'dialog') and self.page.dialog:
                self.page.dialog.open = False
                self.page.dialog = None  # ← Agregar esta línea
                self.page.update()
        except Exception as e:
            logger.error(f"Error cerrando diálogo: {e}")
    
    # ================================
    # MESSAGE/DIALOG FUNCTIONS
    # ================================
    
    def show_error(self, message):
        """Muestra un mensaje de error"""
        self.status_text.value = f"❌ {message}"
        self.status_text.color = ft.Colors.RED_600
        self.page.update()
        
        dialog = ft.AlertDialog(
            title=ft.Text("❌ Error"),
            content=ft.Text(message),
            actions=[ft.TextButton("OK", on_click=lambda e: self.close_dialog())]
        )
        
        self.page.dialog = dialog
        dialog.open = True
        self.page.update()
    
    def show_warning(self, message):
        """Muestra un mensaje de advertencia"""
        self.status_text.value = f"⚠️ {message}"
        self.status_text.color = ft.Colors.ORANGE_600
        self.page.update()
        
        dialog = ft.AlertDialog(
            title=ft.Text("⚠️ Advertencia"),
            content=ft.Text(message),
            actions=[ft.TextButton("OK", on_click=lambda e: self.close_dialog())]
        )
        
        self.page.dialog = dialog
        dialog.open = True
        self.page.update()
    
    def show_info(self, title, message):
        """Muestra un mensaje informativo"""
        dialog = ft.AlertDialog(
            title=ft.Text(f"ℹ️ {title}"),
            content=ft.Container(
                content=ft.Text(message, size=12),
                width=500,
                height=300
            ),
            actions=[ft.TextButton("OK", on_click=lambda e: self.close_dialog())]
        )
        
        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

# ================================
# MAIN APPLICATION ENTRY POINT
# ================================

def main(page: ft.Page):
    """Función principal que inicia la aplicación"""
    try:
        app = R4CSVExporterApp(page)
        logger.info("Aplicación R4 CSV Exporter iniciada exitosamente")
    except Exception as e:
        logger.error(f"Error iniciando aplicación: {e}")
        
        # Mostrar error básico si falla la inicialización
        page.title = "Error - R4 CSV Exporter"
        error_text = ft.Text(
            f"❌ Error iniciando aplicación:\n\n{str(e)}\n\nRevisa los logs para más detalles.",
            size=14,
            color=ft.Colors.RED_600,
            text_align=ft.TextAlign.CENTER
        )
        
        page.add(
            ft.Container(
                content=error_text,
                alignment=ft.alignment.center,
                expand=True
            )
        )

if __name__ == "__main__":
    # Configurar logging adicional para archivo
    try:
        log_file = Path(DEFAULT_CSV_FOLDER) / "r4_exporter.log"
        log_file.parent.mkdir(parents=True, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(file_handler)
        
        logger.info("="*50)
        logger.info("R4 CSV Exporter - Iniciando aplicación")
        logger.info(f"Versión Python: {platform.python_version()}")
        logger.info(f"Sistema: {platform.system()} {platform.release()}")
        logger.info(f"Carpeta CSV por defecto: {DEFAULT_CSV_FOLDER}")
        logger.info(f"BD por defecto: {DEFAULT_DB_PATH}")
        logger.info("="*50)
        
    except Exception as e:
        print(f"Warning: No se pudo configurar logging a archivo: {e}")
    
    try:
        # Verificar dependencias críticas
        required_modules = ['flet', 'pandas', 'sqlite3']
        missing_modules = []
        
        for module in required_modules:
            try:
                __import__(module)
            except ImportError:
                missing_modules.append(module)
        
        if missing_modules:
            print(f"❌ Error: Faltan módulos requeridos: {', '.join(missing_modules)}")
            print("📦 Instala con: pip install flet pandas")
            input("Presiona Enter para salir...")
            exit(1)
        
        # Iniciar aplicación Flet
        print("🚀 Iniciando R4 Database → CSV Exporter...")
        print("📊 Aplicación para exportar datos de R4 a Looker Studio")
        print("🔗 Conectando interfaz gráfica...")
        
        ft.app(
            target=main,
            name="R4 CSV Exporter",
            assets_dir="assets"  # Para futuros iconos/recursos
        )
        
    except KeyboardInterrupt:
        print("\n👋 Aplicación cerrada por el usuario")
        logger.info("Aplicación cerrada por el usuario (Ctrl+C)")
    except Exception as e:
        print(f"❌ Error crítico: {e}")
        logger.error(f"Error crítico en main: {e}")
        input("Presiona Enter para salir...")
    finally:
        logger.info("Aplicación R4 CSV Exporter finalizada")






