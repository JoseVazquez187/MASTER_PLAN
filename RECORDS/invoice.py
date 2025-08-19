# =============================================================================
# MATRIX INVOICE SYSTEM - FLET VERSION
# =============================================================================
# Versión moderna con Flet manteniendo toda la lógica original
# =============================================================================

import flet as ft
import os
import sys
import pandas as pd
import sqlite3
from datetime import datetime
from dataclasses import dataclass
from typing import List, Dict, Any, Optional, Tuple
from abc import ABC, abstractmethod
import asyncio
import threading

# =============================================================================
# CONFIGURACIÓN
# =============================================================================

@dataclass
class DatabaseConfig:
    """Configuración de la base de datos"""
    folder_path: str = r'C:\Users\j.vazquez\Desktop\KITTING PROCESS\invoice'
    db_name: str = 'demand_check.db'
    
    @property
    def connection_path(self) -> str:
        return os.path.join(self.folder_path, self.db_name)

@dataclass
class UIConfig:
    """Configuración de la interfaz de usuario"""
    # Colores Matrix modernos
    matrix_green: str = '#00FF41'
    matrix_dark_green: str = '#00DC2A'
    matrix_black: str = '#0D1117'
    matrix_gray: str = '#21262D'
    matrix_light_gray: str = '#30363D'
    matrix_yellow: str = '#F0C649'
    matrix_red: str = '#F85149'
    matrix_orange: str = '#FB8500'
    matrix_blue: str = '#58A6FF'
    matrix_purple: str = '#BC8CFF'
    
    # Gradientes
    primary_gradient: str = "linear-gradient(135deg, #00FF41 0%, #00DC2A 100%)"
    dark_gradient: str = "linear-gradient(135deg, #0D1117 0%, #21262D 100%)"
    
    # Dimensiones
    window_width: int = 1400
    window_height: int = 900

@dataclass
class DataConfig:
    """Configuración de datos y columnas esperadas"""
    expected_columns: List[str] = None
    
    def __post_init__(self):
        if self.expected_columns is None:
            self.expected_columns = [
                'Entity-Code', 'Invoice-No', 'Line', 'SO-No', 'Customer', 'Type', 'PO-No', 'A/C', 'Spr-C',
                'O-Cd', 'Proj', 'TC', 'Item Number', 'Description', 'UM', 'S-Qty', 'Price', 'Total',
                'CustReqDt', 'Req-Date', 'Pln-Ship', 'Inv-Dt', 'Inv-Time', 'Ord-Dt', 'SODueDt', 'Due-Dt',
                'DG', 'Std Cost', 'Shipped-Dt', 'ShipTo', 'ViaDesc', 'Tracking No', 'Invoice Line Memo',
                'Lot No (Qty)', 'User-Id', 'Cust-Item', 'Buyer-Name', 'Issue-Date'
            ]

class AppSettings:
    """Configuración principal de la aplicación"""
    
    def __init__(self):
        self.database = DatabaseConfig()
        self.ui = UIConfig()
        self.data = DataConfig()
        
        # Configuraciones adicionales
        self.app_title = "Matrix Validation System"
        self.app_subtitle = "Proyecto: INVOICED"
        self.debug_mode = False

# =============================================================================
# GESTOR DE BASE DE DATOS (Mantiene la lógica original)
# =============================================================================

class DatabaseManager:
    """Gestor centralizado de operaciones de base de datos"""
    
    def __init__(self, db_config: DatabaseConfig, data_config: DataConfig):
        self.db_config = db_config
        self.data_config = data_config
        self._ensure_database_exists()
    
    def _ensure_database_exists(self):
        """Asegura que la base de datos y tablas existan"""
        conn = sqlite3.connect(self.db_config.connection_path)
        cursor = conn.cursor()
        
        # Crear tabla de control de carga
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS invoice_loaded (
                last_loaded_date TEXT
            )
        """)
        
        # Crear tabla principal de facturas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS invoiced (
                Entity_Code TEXT, Invoice_No TEXT, Line TEXT, SO_No TEXT, Customer TEXT, Type TEXT,
                PO_No TEXT, A_C TEXT, Spr_C TEXT, O_Cd TEXT, Proj TEXT, TC TEXT, Item_Number TEXT,
                Description TEXT, UM TEXT, S_Qty TEXT, Price TEXT, Total TEXT, CustReqDt TEXT,
                Req_Date TEXT, Pln_Ship TEXT, Inv_Dt TEXT, Inv_Time TEXT, Ord_Dt TEXT, SODueDt TEXT,
                Due_Dt TEXT, DG TEXT, Std_Cost TEXT, Shipped_Dt TEXT, ShipTo TEXT, ViaDesc TEXT,
                Tracking_No TEXT, Invoice_Line_Memo TEXT, Lot_No_Qty TEXT, User_Id TEXT, Cust_Item TEXT,
                Buyer_Name TEXT, Issue_Date TEXT
            )
        """)
        
        conn.commit()
        conn.close()
    
    def get_connection(self) -> sqlite3.Connection:
        """Obtiene una conexión a la base de datos"""
        return sqlite3.connect(self.db_config.connection_path)
    
    def get_last_loaded_date(self) -> Optional[str]:
        """Obtiene la última fecha cargada"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT last_loaded_date FROM invoice_loaded ORDER BY rowid DESC LIMIT 1")
            result = cursor.fetchone()
            return result[0] if result else None
    
    def get_database_stats(self) -> Dict[str, Any]:
        """Obtiene estadísticas de la base de datos"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Verificar si la tabla existe
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='invoiced'")
            table_exists = cursor.fetchone() is not None
            
            if not table_exists:
                return {
                    'table_exists': False,
                    'total_records': 0,
                    'min_date': None,
                    'max_date': None
                }
            
            # Obtener estadísticas
            cursor.execute("SELECT COUNT(*) FROM invoiced")
            total_records = cursor.fetchone()[0]
            
            cursor.execute("SELECT MIN(Inv_Dt), MAX(Inv_Dt) FROM invoiced WHERE Inv_Dt IS NOT NULL AND Inv_Dt != ''")
            date_range = cursor.fetchone()
            
            return {
                'table_exists': True,
                'total_records': total_records,
                'min_date': date_range[0] if date_range[0] else None,
                'max_date': date_range[1] if date_range[1] else None
            }
    
    def clear_invoiced_table(self):
        """Limpia la tabla de facturas"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM invoiced")
            conn.commit()
    
    def insert_dataframe(self, df: pd.DataFrame):
        """Inserta un DataFrame en la tabla de facturas"""
        with self.get_connection() as conn:
            df.to_sql('invoiced', conn, if_exists='append', index=False)
    
    def delete_records_from_date(self, date_str: str) -> int:
        """Elimina registros desde una fecha específica"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM invoiced WHERE Inv_Dt >= ?", (date_str,))
            conn.commit()
            return cursor.rowcount
    
    def update_last_loaded_date(self, date_str: str):
        """Actualiza la última fecha cargada"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM invoice_loaded")
            cursor.execute("INSERT INTO invoice_loaded VALUES (?)", (date_str,))
            conn.commit()

# =============================================================================
# PROCESADOR DE DATOS (Mantiene la lógica original)
# =============================================================================

class DataProcessor:
    """Procesador de archivos de datos de facturas"""
    
    def __init__(self, data_config: DataConfig):
        self.data_config = data_config
    
    def clean_column_names(self, columns: List[str]) -> List[str]:
        """Limpia y normaliza nombres de columnas"""
        cleaned = []
        for col in columns:
            col = col.strip()
            # Mapeo específico de columnas problemáticas
            column_mapping = {
                'A/C': 'A_C',
                'Item Number': 'Item_Number',
                'Std Cost': 'Std_Cost',
                'Tracking No': 'Tracking_No',
                'Invoice Line Memo': 'Invoice_Line_Memo',
                'Lot No (Qty)': 'Lot_No_Qty',
                'User-Id': 'User_Id',
                'Cust-Item': 'Cust_Item',
                'Buyer-Name': 'Buyer_Name',
                'Issue-Date': 'Issue_Date'
            }
            
            if col in column_mapping:
                col = column_mapping[col]
            else:
                # Limpieza general
                col = col.replace(' ', '_').replace('-', '_')
                col = col.replace('(', '').replace(')', '')
                col = col.replace('#', '').replace('/', '_')
                col = col.replace('[', '').replace(']', '')
                col = col.replace('.', '')
            
            cleaned.append(col)
        return cleaned
    
    def read_excel_file(self, file_path: str) -> Tuple[pd.DataFrame, List[str]]:
        """
        Lee un archivo Excel y retorna el DataFrame procesado y una lista de errores
        """
        errors = []
        
        try:
            # Leer archivo
            df = pd.read_excel(file_path, dtype=str)
            df = df.fillna('')
            df = df.applymap(lambda x: x.replace("'", "") if isinstance(x, str) else x)
            
            # Limpiar nombres de columnas
            df.columns = self.clean_column_names(df.columns)
            
            # Verificar columnas esperadas
            expected_clean = self.clean_column_names(self.data_config.expected_columns)
            missing_cols = set(expected_clean) - set(df.columns)
            
            if missing_cols:
                errors.append(f"Columnas faltantes: {missing_cols}")
                return pd.DataFrame(), errors
            
            # Procesar fechas
            df['Inv_Dt'] = pd.to_datetime(df['Inv_Dt'], errors='coerce')
            
            return df[expected_clean], errors
            
        except Exception as e:
            errors.append(f"Error leyendo archivo: {str(e)}")
            return pd.DataFrame(), errors
    
    def get_excel_files_in_directory(self, directory: str) -> List[str]:
        """Obtiene lista de archivos Excel en un directorio"""
        if not os.path.exists(directory):
            return []
        
        return [f for f in os.listdir(directory) 
                if f.endswith('.xlsx') and not f.startswith('~$')]

# =============================================================================
# COMPONENTES UI MODERNOS
# =============================================================================

class ModernCard(ft.Container):
    """Tarjeta moderna con efecto glassmorphism"""
    
    def __init__(self, content, padding=20, gradient=None, blur=10, **kwargs):
        super().__init__(
            content=content,
            padding=padding,
            border_radius=16,
            bgcolor=ft.Colors.with_opacity(0.1, ft.Colors.WHITE),
            border=ft.border.all(1, ft.Colors.with_opacity(0.2, ft.Colors.WHITE)),
            blur=ft.Blur(sigma_x=blur, sigma_y=blur),
            **kwargs
        )

class MatrixButton(ft.ElevatedButton):
    """Botón con estilo Matrix moderno"""
    
    def __init__(self, text, on_click=None, icon=None, color_scheme="primary", **kwargs):
        Colors = {
            "primary": {"bg": "#00FF41", "fg": "#000000"},
            "danger": {"bg": "#F85149", "fg": "#FFFFFF"},
            "warning": {"bg": "#FB8500", "fg": "#FFFFFF"},
            "secondary": {"bg": "#30363D", "fg": "#00FF41"}
        }
        
        scheme = Colors.get(color_scheme, Colors["primary"])
        
        super().__init__(
            text=text,
            on_click=on_click,
            icon=icon,
            style=ft.ButtonStyle(
                bgcolor=scheme["bg"],
                color=scheme["fg"],
                elevation=8,
                shape=ft.RoundedRectangleBorder(radius=12),
                padding=ft.padding.symmetric(horizontal=24, vertical=16),
                text_style=ft.TextStyle(
                    weight=ft.FontWeight.BOLD,
                    size=14
                ),
                animation_duration=200,
                side=ft.BorderSide(2, ft.Colors.with_opacity(0.3, ft.Colors.WHITE))
            ),
            **kwargs
        )

class LogViewer(ft.Container):
    """Visor de logs moderno con scroll automático"""
    
    def __init__(self, height=400):
        self.log_content = ft.Column(
            scroll=ft.ScrollMode.AUTO,
            auto_scroll=True,
            spacing=2
        )
        
        super().__init__(
            content=ft.Container(
                content=self.log_content,
                padding=16,
                bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE),
                border_radius=12,
                border=ft.border.all(1, ft.Colors.with_opacity(0.2, ft.Colors.WHITE))
            ),
            height=height,
            expand=True
        )
    
    def add_log(self, message: str, log_type: str = "info"):
        """Añade un mensaje al log"""
        Colors = {
            "info": "#00FF41",
            "error": "#F85149", 
            "warning": "#FB8500",
            "success": "#00DC2A"
        }
        
        icon_map = {
            "info": ft.Icons.INFO_OUTLINE,
            "error": ft.Icons.ERROR_OUTLINE,
            "warning": ft.Icons.WARNING_AMBER_OUTLINED,
            "success": ft.Icons.CHECK_CIRCLE_OUTLINE
        }
        
        log_entry = ft.Row([
            ft.Icon(
                icon_map.get(log_type, ft.Icons.INFO_OUTLINE),
                color=Colors.get(log_type, "#00FF41"),
                size=16
            ),
            ft.Text(
                f"[{datetime.now().strftime('%H:%M:%S')}] {message}",
                color=Colors.get(log_type, "#00FF41"),
                size=12,
                font_family="Consolas"
            )
        ], spacing=8)
        
        self.log_content.controls.append(log_entry)
        self.update()
    
    def clear_logs(self):
        """Limpia todos los logs"""
        self.log_content.controls.clear()
        self.update()

class StatsCard(ft.Container):
    """Tarjeta de estadísticas moderna"""
    
    def __init__(self, title: str, value: str, icon: str, color: str = "#00FF41"):
        super().__init__(
            content=ft.Column([
                ft.Row([
                    ft.Icon(icon, color=color, size=32),
                    ft.Column([
                        ft.Text(
                            title,
                            size=12,
                            color=ft.Colors.with_opacity(0.7, ft.Colors.WHITE),
                            weight=ft.FontWeight.W_500
                        ),
                        ft.Text(
                            value,
                            size=24,
                            color=color,
                            weight=ft.FontWeight.BOLD
                        )
                    ], spacing=2, expand=True)
                ], spacing=16)
            ]),
            padding=20,
            border_radius=16,
            bgcolor=ft.Colors.with_opacity(0.1, ft.Colors.WHITE),
            border=ft.border.all(1, ft.Colors.with_opacity(0.2, ft.Colors.WHITE)),
            expand=True
        )

# =============================================================================
# PESTAÑAS MODERNAS
# =============================================================================

class BaseTab(ABC):
    """Clase base para todas las pestañas"""
    
    def __init__(self, db_manager: DatabaseManager, data_processor: DataProcessor, 
                 on_refresh_callback=None):
        self.db_manager = db_manager
        self.data_processor = data_processor
        self.on_refresh_callback = on_refresh_callback
        
        # Crear contenido de la pestaña
        self.content = self.build_content()
    
    @abstractmethod
    def build_content(self) -> ft.Control:
        """Construye el contenido de la pestaña"""
        pass
    
    @abstractmethod
    def get_tab_title(self) -> str:
        """Retorna el título de la pestaña"""
        pass
    
    @abstractmethod
    def get_tab_icon(self) -> str:
        """Retorna el icono de la pestaña"""
        pass

class DashboardTab(BaseTab):
    """Pestaña principal con información del sistema"""
    
    def get_tab_title(self) -> str:
        return "Dashboard"
    
    def get_tab_icon(self) -> str:
        return ft.Icons.DASHBOARD_OUTLINED
    
    def build_content(self) -> ft.Control:
        # Estadísticas principales
        self.stats_row = ft.Row(spacing=16, expand=True)
        
        # Información de archivos
        self.files_list = ft.Column(spacing=8, scroll=ft.ScrollMode.AUTO)
        
        # Construir layout
        content = ft.Column([
            # Header
            ft.Container(
                content=ft.Column([
                    ft.Text(
                        "Sistema de Gestión de Facturas",
                        size=28,
                        weight=ft.FontWeight.BOLD,
                        color="#00FF41"
                    ),
                    ft.Text(
                        "Monitoreo y control de datos INVOICED",
                        size=16,
                        color=ft.Colors.with_opacity(0.7, ft.Colors.WHITE)
                    )
                ]),
                padding=20
            ),
            
            # Estadísticas
            ModernCard(
                content=ft.Column([
                    ft.Text(
                        "Estadísticas de Base de Datos",
                        size=18,
                        weight=ft.FontWeight.BOLD,
                        color="#00FF41"
                    ),
                    self.stats_row
                ], spacing=16),
                padding=20
            ),
            
            # Archivos disponibles
            ModernCard(
                content=ft.Column([
                    ft.Row([
                        ft.Text(
                            "Archivos Disponibles",
                            size=18,
                            weight=ft.FontWeight.BOLD,
                            color="#00FF41"
                        ),
                        ft.IconButton(
                            icon=ft.Icons.REFRESH,
                            icon_color="#00FF41",
                            on_click=self.refresh_data,
                            tooltip="Actualizar información"
                        )
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    ft.Container(
                        content=self.files_list,
                        height=200,
                        border_radius=8,
                        bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE),
                        padding=16
                    )
                ], spacing=16),
                padding=20,
                expand=True
            )
        ], spacing=16, expand=True)
        
        # Cargar datos iniciales
        self.refresh_data(None)
        
        return ft.Container(
            content=content,
            padding=20,
            expand=True
        )
    
    def refresh_data(self, e):
        """Actualiza la información del dashboard"""
        try:
            # Obtener estadísticas
            stats = self.db_manager.get_database_stats()
            
            # Limpiar estadísticas anteriores
            self.stats_row.controls.clear()
            
            # Crear tarjetas de estadísticas
            total_records = f"{stats['total_records']:,}" if stats['table_exists'] else "0"
            
            self.stats_row.controls.extend([
                StatsCard(
                    "Total Registros",
                    total_records,
                    ft.Icons.STORAGE,
                    "#58A6FF"
                ),
                StatsCard(
                    "Fecha Mínima",
                    stats['min_date'] or "N/A",
                    ft.Icons.CALENDAR_TODAY,
                    "#BC8CFF"
                ),
                StatsCard(
                    "Fecha Máxima",
                    stats['max_date'] or "N/A",
                    ft.Icons.EVENT,
                    "#F0C649"
                ),
                StatsCard(
                    "Estado BD",
                    "Activa" if stats['table_exists'] else "Vacía",
                    ft.Icons.DATABASE,
                    "#00DC2A" if stats['table_exists'] else "#F85149"
                )
            ])
            
            # Obtener archivos
            files = self.data_processor.get_excel_files_in_directory(
                self.db_manager.db_config.folder_path
            )
            
            # Limpiar lista de archivos
            self.files_list.controls.clear()
            
            if files:
                for i, file in enumerate(sorted(files), 1):
                    file_item = ft.Container(
                        content=ft.Row([
                            ft.Icon(ft.Icons.DESCRIPTION, color="#00FF41", size=20),
                            ft.Text(
                                f"{i:02d}. {file}",
                                size=14,
                                color="#FFFFFF",
                                font_family="Consolas"
                            )
                        ], spacing=12),
                        padding=ft.padding.symmetric(vertical=4, horizontal=8),
                        border_radius=6,
                        bgcolor=ft.Colors.with_opacity(0.1, ft.Colors.WHITE) if i % 2 == 0 else None
                    )
                    self.files_list.controls.append(file_item)
            else:
                self.files_list.controls.append(
                    ft.Container(
                        content=ft.Row([
                            ft.Icon(ft.Icons.WARNING_AMBER, color="#FB8500", size=20),
                            ft.Text(
                                "No se encontraron archivos .xlsx",
                                size=14,
                                color="#FB8500"
                            )
                        ], spacing=12),
                        padding=12
                    )
                )
            
            # Actualizar UI
            self.stats_row.update()
            self.files_list.update()
            
        except Exception as ex:
            print(f"Error actualizando dashboard: {ex}")

class FullLoadTab(BaseTab):
    """Pestaña para carga completa de datos"""
    
    def get_tab_title(self) -> str:
        return "Carga Completa"
    
    def get_tab_icon(self) -> str:
        return ft.Icons.REFRESH
    
    def build_content(self) -> ft.Control:
        # Log viewer
        self.log_viewer = LogViewer(height=400)
        
        # Progress bar
        self.progress_bar = ft.ProgressBar(
            visible=False,
            color="#00FF41",
            bgcolor=ft.Colors.with_opacity(0.2, ft.Colors.WHITE)
        )
        
        # Botón de carga
        self.load_button = MatrixButton(
            text="Iniciar Carga Completa",
            icon=ft.Icons.REFRESH,
            color_scheme="danger",
            on_click=self.execute_full_load
        )
        
        content = ft.Column([
            # Header
            ft.Container(
                content=ft.Column([
                    ft.Text(
                        "Carga Completa de Datos",
                        size=24,
                        weight=ft.FontWeight.BOLD,
                        color="#00FF41"
                    ),
                    ft.Text(
                        "Elimina todos los registros existentes y carga todos los archivos .xlsx",
                        size=14,
                        color=ft.Colors.with_opacity(0.7, ft.Colors.WHITE)
                    )
                ]),
                padding=20
            ),
            
            # Warning card
            ModernCard(
                content=ft.Row([
                    ft.Icon(ft.Icons.WARNING_AMBER, color="#FB8500", size=32),
                    ft.Column([
                        ft.Text(
                            "Advertencia",
                            size=16,
                            weight=ft.FontWeight.BOLD,
                            color="#FB8500"
                        ),
                        ft.Text(
                            "Esta operación eliminará todos los datos existentes de la base de datos.",
                            size=12,
                            color=ft.Colors.with_opacity(0.8, ft.Colors.WHITE)
                        )
                    ], expand=True, spacing=4)
                ], spacing=16),
                padding=20
            ),
            
            # Progress
            ft.Container(
                content=self.progress_bar,
                padding=ft.padding.symmetric(horizontal=20)
            ),
            
            # Log viewer
            ModernCard(
                content=ft.Column([
                    ft.Row([
                        ft.Text(
                            "Log de Operaciones",
                            size=16,
                            weight=ft.FontWeight.BOLD,
                            color="#00FF41"
                        ),
                        ft.IconButton(
                            icon=ft.Icons.CLEAR,
                            icon_color="#F85149",
                            on_click=lambda e: self.log_viewer.clear_logs(),
                            tooltip="Limpiar log"
                        )
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    self.log_viewer
                ], spacing=12),
                padding=20,
                expand=True
            ),
            
            # Action button
            ft.Container(
                content=self.load_button,
                padding=20,
                alignment=ft.alignment.center
            )
        ], spacing=16, expand=True)
        
        return ft.Container(
            content=content,
            padding=20,
            expand=True
        )
    
    def execute_full_load(self, e):
        """Ejecuta la carga completa en un hilo separado"""
        self.load_button.disabled = True
        self.progress_bar.visible = True
        self.progress_bar.value = None  # Indeterminate
        self.load_button.update()
        self.progress_bar.update()
        
        # Ejecutar en hilo separado para no bloquear UI
        threading.Thread(target=self._full_load_worker, daemon=True).start()
    
    def _full_load_worker(self):
        """Worker para la carga completa"""
        try:
            self.log_viewer.add_log("Iniciando carga completa de datos...", "info")
            
            # Limpiar tabla
            self.db_manager.clear_invoiced_table()
            self.log_viewer.add_log("Base de datos limpiada exitosamente", "success")
            
            # Obtener archivos
            files = self.data_processor.get_excel_files_in_directory(
                self.db_manager.db_config.folder_path
            )
            
            if not files:
                self.log_viewer.add_log("No se encontraron archivos .xlsx", "error")
                return
            
            self.log_viewer.add_log(f"Encontrados {len(files)} archivos para procesar", "info")
            
            all_data = pd.DataFrame()
            total_inserted = 0
            
            # Procesar cada archivo
            for i, file in enumerate(sorted(files), 1):
                self.log_viewer.add_log(f"Procesando archivo {i}/{len(files)}: {file}", "info")
                
                file_path = os.path.join(self.db_manager.db_config.folder_path, file)
                df, errors = self.data_processor.read_excel_file(file_path)
                
                if errors:
                    for error in errors:
                        self.log_viewer.add_log(f"Error en {file}: {error}", "error")
                    continue
                
                if not df.empty:
                    try:
                        # Preparar datos para insertar
                        df_copy = df.copy()
                        df_copy['Inv_Dt'] = df_copy['Inv_Dt'].dt.strftime('%Y-%m-%d')
                        
                        # Insertar en base de datos
                        self.db_manager.insert_dataframe(df_copy)
                        
                        self.log_viewer.add_log(f"✓ {file}: {len(df_copy):,} registros insertados", "success")
                        total_inserted += len(df_copy)
                        all_data = pd.concat([all_data, df], ignore_index=True)
                        
                    except Exception as ex:
                        self.log_viewer.add_log(f"Error insertando {file}: {ex}", "error")
            
            # Actualizar fecha de última carga
            if total_inserted > 0 and not all_data.empty:
                max_date = all_data['Inv_Dt'].max().strftime('%Y-%m-%d')
                self.db_manager.update_last_loaded_date(max_date)
                self.log_viewer.add_log(f"Fecha máxima actualizada: {max_date}", "info")
            
            self.log_viewer.add_log(f"Carga completa finalizada: {total_inserted:,} registros totales", "success")
            
            # Notificar callback de actualización
            if self.on_refresh_callback:
                self.on_refresh_callback()
                
        except Exception as ex:
            self.log_viewer.add_log(f"Error durante carga completa: {ex}", "error")
        finally:
            # Restaurar UI
            self.load_button.disabled = False
            self.progress_bar.visible = False
            self.load_button.update()
            self.progress_bar.update()

class AppendTab(BaseTab):
    """Pestaña para carga incremental (append)"""
    
    def get_tab_title(self) -> str:
        return "Carga Incremental"
    
    def get_tab_icon(self) -> str:
        return ft.Icons.ADD_CIRCLE_OUTLINE
    
    def build_content(self) -> ft.Control:
        # File picker result
        def pick_files_result(e: ft.FilePickerResultEvent):
            if e.files:
                selected_file = e.files[0]
                self.selected_file_path = selected_file.path
                self.file_name_text.value = selected_file.name
                self.append_button.disabled = False
                self.file_name_text.update()
                self.append_button.update()
                self.log_viewer.add_log(f"Archivo seleccionado: {selected_file.name}", "info")
        
        # File picker
        self.file_picker = ft.FilePicker(on_result=pick_files_result)
        
        # Selected file display
        self.file_name_text = ft.Text(
            "Ningún archivo seleccionado",
            size=14,
            color="#FB8500",
            weight=ft.FontWeight.W_500
        )
        
        # Log viewer
        self.log_viewer = LogViewer(height=350)
        
        # Progress bar
        self.progress_bar = ft.ProgressBar(
            visible=False,
            color="#00FF41",
            bgcolor=ft.Colors.with_opacity(0.2, ft.Colors.WHITE)
        )
        
        # Buttons
        self.select_button = MatrixButton(
            text="Seleccionar Archivo",
            icon=ft.Icons.FOLDER_OPEN,
            color_scheme="secondary",
            on_click=lambda e: self.file_picker.pick_files(
                dialog_title="Seleccionar archivo Excel",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx"]
            )
        )
        
        self.append_button = MatrixButton(
            text="Ejecutar Append",
            icon=ft.Icons.PLAY_ARROW,
            color_scheme="warning",
            on_click=self.execute_append,
            disabled=True
        )
        
        content = ft.Column([
            # File picker overlay
            self.file_picker,
            
            # Header
            ft.Container(
                content=ft.Column([
                    ft.Text(
                        "Carga Incremental (Append)",
                        size=24,
                        weight=ft.FontWeight.BOLD,
                        color="#00FF41"
                    ),
                    ft.Text(
                        "Carga datos de un archivo específico manejando duplicados por fecha",
                        size=14,
                        color=ft.Colors.with_opacity(0.7, ft.Colors.WHITE)
                    )
                ]),
                padding=20
            ),
            
            # File selection card
            ModernCard(
                content=ft.Column([
                    ft.Text(
                        "Selección de Archivo",
                        size=16,
                        weight=ft.FontWeight.BOLD,
                        color="#00FF41"
                    ),
                    ft.Row([
                        ft.Icon(ft.Icons.DESCRIPTION, color="#58A6FF", size=24),
                        ft.Column([
                            ft.Text("Archivo:", size=12, color=ft.Colors.with_opacity(0.7, ft.Colors.WHITE)),
                            self.file_name_text
                        ], expand=True, spacing=4)
                    ], spacing=16),
                    ft.Row([
                        self.select_button,
                        self.append_button
                    ], spacing=16, alignment=ft.MainAxisAlignment.CENTER)
                ], spacing=16),
                padding=20
            ),
            
            # Progress
            ft.Container(
                content=self.progress_bar,
                padding=ft.padding.symmetric(horizontal=20)
            ),
            
            # Log viewer
            ModernCard(
                content=ft.Column([
                    ft.Row([
                        ft.Text(
                            "Log de Operaciones",
                            size=16,
                            weight=ft.FontWeight.BOLD,
                            color="#00FF41"
                        ),
                        ft.IconButton(
                            icon=ft.Icons.CLEAR,
                            icon_color="#F85149",
                            on_click=lambda e: self.log_viewer.clear_logs(),
                            tooltip="Limpiar log"
                        )
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    self.log_viewer
                ], spacing=12),
                padding=20,
                expand=True
            )
        ], spacing=16, expand=True)
        
        return ft.Container(
            content=content,
            padding=20,
            expand=True
        )
    
    def execute_append(self, e):
        """Ejecuta el proceso de append"""
        if not hasattr(self, 'selected_file_path'):
            self.log_viewer.add_log("No hay archivo seleccionado", "error")
            return
        
        self.append_button.disabled = True
        self.select_button.disabled = True
        self.progress_bar.visible = True
        self.progress_bar.value = None
        self.append_button.update()
        self.select_button.update()
        self.progress_bar.update()
        
        # Ejecutar en hilo separado
        threading.Thread(target=self._append_worker, daemon=True).start()
    
    def _append_worker(self):
        """Worker para el proceso append"""
        try:
            self.log_viewer.add_log("Iniciando proceso de append...", "info")
            
            # Leer archivo seleccionado
            self.log_viewer.add_log("Leyendo archivo seleccionado...", "info")
            df_new, errors = self.data_processor.read_excel_file(self.selected_file_path)
            
            if errors:
                for error in errors:
                    self.log_viewer.add_log(f"Error: {error}", "error")
                return
            
            if df_new.empty:
                self.log_viewer.add_log("El archivo está vacío o no se pudo leer", "error")
                return
            
            # Obtener fechas del archivo
            min_date_new = df_new['Inv_Dt'].min()
            max_date_new = df_new['Inv_Dt'].max()
            
            self.log_viewer.add_log(f"Rango de fechas: {min_date_new.strftime('%Y-%m-%d')} a {max_date_new.strftime('%Y-%m-%d')}", "info")
            self.log_viewer.add_log(f"Total de registros: {len(df_new):,}", "info")
            
            # Verificar última fecha en BD
            last_loaded = self.db_manager.get_last_loaded_date()
            
            if last_loaded:
                last_date_db = datetime.strptime(last_loaded, '%Y-%m-%d')
                self.log_viewer.add_log(f"Última fecha en BD: {last_date_db.strftime('%Y-%m-%d')}", "info")
                
                # Verificar si hay overlap
                if min_date_new <= last_date_db:
                    self.log_viewer.add_log(f"Detectado overlap. Eliminando registros desde {min_date_new.strftime('%Y-%m-%d')}", "warning")
                    
                    # Eliminar registros desde la fecha mínima del archivo
                    deleted_count = self.db_manager.delete_records_from_date(min_date_new.strftime('%Y-%m-%d'))
                    self.log_viewer.add_log(f"Eliminados {deleted_count:,} registros", "info")
                else:
                    self.log_viewer.add_log("No hay overlap de fechas", "success")
            else:
                self.log_viewer.add_log("No hay fecha previa en BD", "info")
            
            # Insertar nuevos datos
            self.log_viewer.add_log("Insertando nuevos registros...", "info")
            df_copy = df_new.copy()
            df_copy['Inv_Dt'] = df_copy['Inv_Dt'].dt.strftime('%Y-%m-%d')
            self.db_manager.insert_dataframe(df_copy)
            
            # Actualizar fecha de última carga
            self.db_manager.update_last_loaded_date(max_date_new.strftime('%Y-%m-%d'))
            
            self.log_viewer.add_log(f"Append completado exitosamente", "success")
            self.log_viewer.add_log(f"Registros insertados: {len(df_new):,}", "success")
            self.log_viewer.add_log(f"Nueva fecha máxima: {max_date_new.strftime('%Y-%m-%d')}", "info")
            
            # Notificar callback de actualización
            if self.on_refresh_callback:
                self.on_refresh_callback()
                
        except Exception as ex:
            self.log_viewer.add_log(f"Error durante append: {ex}", "error")
        finally:
            # Restaurar UI
            self.append_button.disabled = False
            self.select_button.disabled = False
            self.progress_bar.visible = False
            self.append_button.update()
            self.select_button.update()
            self.progress_bar.update()

# =============================================================================
# APLICACIÓN PRINCIPAL FLET
# =============================================================================

class MatrixInvoiceApp:
    """Aplicación principal Matrix Invoice System con Flet"""
    
    def __init__(self, page: ft.Page):
        self.page = page
        self.settings = AppSettings()
        
        # Inicializar componentes core
        self.db_manager = DatabaseManager(self.settings.database, self.settings.data)
        self.data_processor = DataProcessor(self.settings.data)
        
        # Configurar página
        self.setup_page()
        
        # Crear interfaz
        self.create_interface()
    
    def setup_page(self):
        """Configura la página principal"""
        self.page.title = self.settings.app_title
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.window_width = self.settings.ui.window_width
        self.page.window_height = self.settings.ui.window_height
        self.page.window_center()
        self.page.bgcolor = self.settings.ui.matrix_black
        self.page.window_maximizable = True
        
        # Configurar tema personalizado
        self.page.theme = ft.Theme(
            color_scheme=ft.Colorscheme(
                primary="#00FF41",
                on_primary="#000000",
                secondary="#30363D",
                on_secondary="#00FF41",
                surface="#0D1117",
                on_surface="#FFFFFF",
                background="#0D1117",
                on_background="#FFFFFF"
            )
        )
    
    def create_interface(self):
        """Crea la interfaz principal"""
        # Header con información de estado
        self.create_header()
        
        # Crear pestañas
        self.create_tabs()
        
        # Layout principal
        main_content = ft.Column([
            self.header_container,
            ft.Container(
                content=self.tabs_container,
                expand=True,
                border_radius=16,
                bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE),
                border=ft.border.all(1, ft.Colors.with_opacity(0.1, ft.Colors.WHITE))
            )
        ], spacing=16, expand=True)
        
        self.page.add(
            ft.Container(
                content=main_content,
                padding=20,
                expand=True
            )
        )
    
    def create_header(self):
        """Crea el header de la aplicación"""
        # Estado de la última fecha
        self.last_date_display = ft.Text(
            "Cargando...",
            size=14,
            color="#F0C649",
            weight=ft.FontWeight.BOLD
        )
        
        # Header principal
        self.header_container = ft.Container(
            content=ft.Column([
                # Título principal
                ft.Row([
                    ft.Icon(ft.Icons.ANALYTICS, color="#00FF41", size=32),
                    ft.Column([
                        ft.Text(
                            self.settings.app_title,
                            size=32,
                            weight=ft.FontWeight.BOLD,
                            color="#00FF41"
                        ),
                        ft.Text(
                            self.settings.app_subtitle,
                            size=16,
                            color=ft.Colors.with_opacity(0.7, ft.Colors.WHITE)
                        )
                    ], spacing=4, expand=True)
                ], spacing=16),
                
                # Barra de estado
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.SCHEDULE, color="#58A6FF", size=20),
                        ft.Text("Última fecha cargada:", size=14, color="#FFFFFF"),
                        self.last_date_display,
                        ft.Container(expand=True),  # Spacer
                        ft.IconButton(
                            icon=ft.Icons.REFRESH,
                            icon_color="#00FF41",
                            on_click=self.refresh_status,
                            tooltip="Actualizar estado"
                        )
                    ], spacing=12),
                    padding=16,
                    border_radius=12,
                    bgcolor=ft.Colors.with_opacity(0.1, ft.Colors.WHITE),
                    border=ft.border.all(1, ft.Colors.with_opacity(0.2, ft.Colors.WHITE))
                )
            ], spacing=20),
            padding=20,
            border_radius=16,
            bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.WHITE),
            border=ft.border.all(1, ft.Colors.with_opacity(0.1, ft.Colors.WHITE))
        )
        
        # Actualizar estado inicial
        self.refresh_status(None)
    
    def create_tabs(self):
        """Crea las pestañas de la aplicación"""
        # Crear instancias de pestañas
        self.dashboard_tab = DashboardTab(self.db_manager, self.data_processor, self.refresh_all_tabs)
        self.full_load_tab = FullLoadTab(self.db_manager, self.data_processor, self.refresh_all_tabs)
        self.append_tab = AppendTab(self.db_manager, self.data_processor, self.refresh_all_tabs)
        
        # Crear tabs
        self.tabs_container = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tab_alignment=ft.TabAlignment.START,
            tabs=[
                ft.Tab(
                    text=self.dashboard_tab.get_tab_title(),
                    icon=self.dashboard_tab.get_tab_icon(),
                    content=self.dashboard_tab.content
                ),
                ft.Tab(
                    text=self.full_load_tab.get_tab_title(),
                    icon=self.full_load_tab.get_tab_icon(),
                    content=self.full_load_tab.content
                ),
                ft.Tab(
                    text=self.append_tab.get_tab_title(),
                    icon=self.append_tab.get_tab_icon(),
                    content=self.append_tab.content
                )
            ],
            expand=True
        )
    
    def refresh_status(self, e):
        """Actualiza el estado en el header"""
        try:
            last_date = self.db_manager.get_last_loaded_date()
            
            if last_date:
                self.last_date_display.value = last_date
                self.last_date_display.color = "#00DC2A"
            else:
                # Buscar fecha máxima en los datos existentes
                stats = self.db_manager.get_database_stats()
                if stats['table_exists'] and stats['max_date']:
                    self.db_manager.update_last_loaded_date(stats['max_date'])
                    self.last_date_display.value = stats['max_date']
                    self.last_date_display.color = "#00DC2A"
                else:
                    self.last_date_display.value = "Sin datos"
                    self.last_date_display.color = "#FB8500"
        except Exception as ex:
            self.last_date_display.value = f"Error: {str(ex)[:20]}..."
            self.last_date_display.color = "#F85149"
        
        self.last_date_display.update()
    
    def refresh_all_tabs(self):
        """Refresca la información en todas las pestañas"""
        self.refresh_status(None)
        # El dashboard se actualizará automáticamente cuando se haga clic en su botón refresh

# =============================================================================
# PUNTO DE ENTRADA PRINCIPAL
# =============================================================================

def main(page: ft.Page):
    """Punto de entrada principal de la aplicación Flet"""
    try:
        # Inicializar la aplicación
        app = MatrixInvoiceApp(page)
        
    except Exception as e:
        # Mostrar error en una página de error
        page.title = "Error - Matrix Invoice System"
        page.bgcolor = "#0D1117"
        
        error_content = ft.Column([
            ft.Icon(ft.Icons.ERROR, color="#F85149", size=64),
            ft.Text(
                "Error Fatal",
                size=24,
                weight=ft.FontWeight.BOLD,
                color="#F85149"
            ),
            ft.Text(
                f"No se pudo iniciar la aplicación:\n{str(e)}",
                size=14,
                color="#FFFFFF",
                text_align=ft.TextAlign.CENTER
            ),
            ft.ElevatedButton(
                text="Cerrar",
                on_click=lambda e: page.window_close(),
                bgcolor="#F85149",
                color="#FFFFFF"
            )
        ], 
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=20)
        
        page.add(
            ft.Container(
                content=error_content,
                alignment=ft.alignment.center,
                expand=True,
                padding=40
            )
        )

if __name__ == "__main__":
    # Ejecutar aplicación Flet
    ft.app(target=main, view=ft.WEB_BROWSER)  # Cambiar a ft.DESKTOP_APP para app desktop

# =============================================================================
# INSTRUCCIONES DE USO PARA LA VERSIÓN FLET:
# =============================================================================
"""
INSTRUCCIONES PARA USAR LA APLICACIÓN FLET MODERNA:

1. INSTALACIÓN:
   pip install flet pandas openpyxl

2. EJECUCIÓN:
   python main.py

3. CARACTERÍSTICAS NUEVAS:
   - Interfaz moderna con efectos glassmorphism
   - Tema Matrix actualizado con gradientes
   - Componentes responsivos y animados
   - Progress bars y feedback visual mejorado
   - Logs en tiempo real con iconos y colores
   - Tarjetas de estadísticas interactivas
   - File picker integrado
   - Threading para operaciones pesadas sin bloquear UI

4. CONFIGURACIÓN:
   - Modifica DatabaseConfig.folder_path para tu directorio
   - Ajusta UIConfig para personalizar colores
   - Cambia ft.WEB_BROWSER por ft.DESKTOP_APP para app de escritorio

5. AGREGAR NUEVAS PESTAÑAS:
   
   class NewTab(BaseTab):
       def get_tab_title(self) -> str:
           return "Nueva Pestaña"
       
       def get_tab_icon(self) -> str:
           return ft.Icons.NEW_RELEASES
       
       def build_content(self) -> ft.Control:
           return ft.Container(
               content=ft.Text("Contenido de nueva pestaña"),
               padding=20
           )
   
   Luego agrégala en create_tabs() de MatrixInvoiceApp

6. PERSONALIZACIÓN:
   - Modifica los colores en UIConfig
   - Ajusta los componentes ModernCard, MatrixButton, etc.
   - Cambia el tema en setup_page()

7. DEPLOYMENT:
   - Para web: flet pack --web
   - Para desktop: flet pack --icon icon.ico
   - Para mobile: flet build apk (requiere configuración adicional)

MEJORAS IMPLEMENTADAS:
- UI completamente rediseñada con Flet
- Componentes modernos reutilizables
- Mejor manejo de errores y feedback
- Threading para operaciones largas
- Logs estructurados con iconos
- Interfaz responsiva
- Animaciones suaves
- Tema Matrix moderno y profesional
"""