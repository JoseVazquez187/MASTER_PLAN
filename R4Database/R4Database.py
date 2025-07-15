#!/usr/bin/env python3
"""
R4 Easy - Sistema Simple de Base de Datos
Versión con actualización masiva de todas las tablas
"""

import sys
import sqlite3
import pandas as pd
import os
import hashlib
import json
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                            QWidget, QPushButton, QLabel, QMessageBox, QProgressBar,
                            QTextEdit, QGroupBox, QComboBox, QTableWidget, QTableWidgetItem,
                            QTabWidget)
from PyQt5.QtCore import QTimer, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont

# Importar configuración
try:
    from config_tables import (TABLES_CONFIG, BASE_PATHS, get_all_tables, 
                            read_file_data, detect_column_format, 
                            process_dataframe_columns)
    print(f"✅ Configuración cargada: {len(TABLES_CONFIG)} tablas disponibles")
except ImportError as e:
    print(f"❌ Error importando configuración: {e}")
    sys.exit(1)

class SimpleUpdater(QThread):
    """Actualizador simple para una tabla"""
    progress_update = pyqtSignal(str)
    finished_update = pyqtSignal(bool, str, dict)
    
    def __init__(self, table_name):
        super().__init__()
        self.table_name = table_name
        self.db_path = os.path.join(BASE_PATHS["db_folder"], "R4Database.db")
        self.tracking_file = BASE_PATHS["tracking_file"]
        
    def run(self):
        try:
            success, message, stats = self.update_table()
            self.finished_update.emit(success, message, stats)
        except Exception as e:
            self.finished_update.emit(False, f"Error: {str(e)}", {})
    
    def get_file_hash(self, file_path):
        """Calcula hash MD5 del archivo"""
        if not os.path.exists(file_path):
            return ""
        hash_md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    
    def load_tracking(self):
        """Carga datos de seguimiento"""
        try:
            if os.path.exists(self.tracking_file):
                with open(self.tracking_file, 'r') as f:
                    return json.load(f)
            return {}
        except:
            return {}
    
    def save_tracking(self, data):
        """Guarda datos de seguimiento"""
        os.makedirs(os.path.dirname(self.tracking_file), exist_ok=True)
        with open(self.tracking_file, 'w') as f:
            json.dump(data, f, indent=2)
    
    def file_has_changed(self, config):
        """Verifica si el archivo ha cambiado"""
        file_path = config["source_file"]
        if not os.path.exists(file_path):
            return False, "Archivo no encontrado"
        
        tracking_data = self.load_tracking()
        table_data = tracking_data.get(self.table_name, {})
        
        current_hash = self.get_file_hash(file_path)
        stored_hash = table_data.get("last_hash", "")
        
        current_modified = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
        stored_modified = table_data.get("last_modified", "")
        
        has_changed = current_hash != stored_hash or current_modified != stored_modified
        
        return has_changed, {
            "current_hash": current_hash,
            "current_modified": current_modified,
            "stored_hash": stored_hash,
            "stored_modified": stored_modified
        }
    
    def update_table(self):
        """Actualiza la tabla"""
        config = TABLES_CONFIG.get(self.table_name)
        if not config:
            return False, f"Configuración no encontrada para {self.table_name}", {}
        
        self.progress_update.emit(f"Verificando cambios en {self.table_name}...")
        
        has_changed, change_info = self.file_has_changed(config)
        
        if not has_changed and isinstance(change_info, dict):
            return False, "No se detectaron cambios en el archivo", {
                "last_updated": change_info.get("stored_modified", "N/A"),
                "records_count": self.load_tracking().get(self.table_name, {}).get("records_count", 0)
            }
        
        if isinstance(change_info, str):  # Error message
            return False, change_info, {}
        
        self.progress_update.emit(f"Leyendo archivo {os.path.basename(config['source_file'])}...")
        
        # Leer archivo
        df = read_file_data(config["source_file"], config)
        
        self.progress_update.emit("Detectando formato de columnas...")
        detected_format = detect_column_format(config["source_file"], config)
        
        self.progress_update.emit("Procesando datos...")
        df_processed = process_dataframe_columns(df, config, detected_format)
        
        self.progress_update.emit("Actualizando base de datos...")
        
        # Conectar a BD
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # Crear tabla específica
            cursor.execute(config["create_table_sql"])
            
            # Limpiar datos anteriores
            cursor.execute(f"DELETE FROM {config['table_name']}")
            
            # Insertar nuevos datos
            df_processed.to_sql(config['table_name'], conn, if_exists='append', index=False)
            
            conn.commit()
            
            # Actualizar tracking
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            tracking_data = self.load_tracking()
            tracking_data[self.table_name] = {
                "file_path": config["source_file"],
                "last_hash": change_info["current_hash"],
                "last_modified": change_info["current_modified"],
                "last_updated": now,
                "records_count": len(df_processed),
                "detected_format": detected_format
            }
            self.save_tracking(tracking_data)
            
            self.progress_update.emit(f"✅ {self.table_name} actualizada exitosamente!")
            
            return True, f"Tabla {self.table_name} actualizada exitosamente", {
                "records_count": len(df_processed),
                "last_updated": now,
                "detected_format": detected_format
            }
            
        finally:
            conn.close()

class BulkUpdater(QThread):
    """Actualizador masivo para todas las tablas"""
    progress_update = pyqtSignal(str)
    table_finished = pyqtSignal(str, bool, str, dict)  # table_name, success, message, stats
    all_finished = pyqtSignal(dict)  # resumen final
    
    def __init__(self, force_update=False):
        super().__init__()
        self.force_update = force_update
        self.db_path = os.path.join(BASE_PATHS["db_folder"], "R4Database.db")
        self.tracking_file = BASE_PATHS["tracking_file"]
        self.results = {}
        
    def run(self):
        try:
            self.update_all_tables()
        except Exception as e:
            self.progress_update.emit(f"❌ Error crítico: {str(e)}")
            self.all_finished.emit({"error": str(e)})
    
    def get_file_hash(self, file_path):
        """Calcula hash MD5 del archivo"""
        if not os.path.exists(file_path):
            return ""
        hash_md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    
    def load_tracking(self):
        """Carga datos de seguimiento"""
        try:
            if os.path.exists(self.tracking_file):
                with open(self.tracking_file, 'r') as f:
                    return json.load(f)
            return {}
        except:
            return {}
    
    def save_tracking(self, data):
        """Guarda datos de seguimiento"""
        os.makedirs(os.path.dirname(self.tracking_file), exist_ok=True)
        with open(self.tracking_file, 'w') as f:
            json.dump(data, f, indent=2)
    
    def update_all_tables(self):
        """Actualiza todas las tablas configuradas"""
        total_tables = len(TABLES_CONFIG)
        updated_count = 0
        skipped_count = 0
        error_count = 0
        
        self.progress_update.emit(f"🚀 Iniciando actualización masiva de {total_tables} tablas...")
        
        # Si es forzar actualización, limpiar tracking
        if self.force_update:
            self.progress_update.emit("⚡ Forzando actualización - limpiando tracking...")
            tracking_data = self.load_tracking()
            for table_name in TABLES_CONFIG.keys():
                if table_name in tracking_data:
                    tracking_data[table_name]["last_hash"] = ""
            self.save_tracking(tracking_data)
        
        # Procesar cada tabla
        for i, (table_name, config) in enumerate(TABLES_CONFIG.items(), 1):
            self.progress_update.emit(f"📋 [{i}/{total_tables}] Procesando tabla: {table_name}")
            
            try:
                success, message, stats = self.update_single_table(table_name, config)
                
                if success:
                    updated_count += 1
                    self.progress_update.emit(f"✅ {table_name}: {message}")
                else:
                    if "No se detectaron cambios" in message:
                        skipped_count += 1
                        self.progress_update.emit(f"ℹ️ {table_name}: Sin cambios")
                    else:
                        error_count += 1
                        self.progress_update.emit(f"❌ {table_name}: {message}")
                
                self.results[table_name] = {
                    "success": success,
                    "message": message,
                    "stats": stats,
                    "status": "updated" if success else ("skipped" if "No se detectaron cambios" in message else "error")
                }
                
                # Emitir señal para tabla individual
                self.table_finished.emit(table_name, success, message, stats)
                
            except Exception as e:
                error_count += 1
                error_msg = f"Error procesando: {str(e)}"
                self.progress_update.emit(f"❌ {table_name}: {error_msg}")
                
                self.results[table_name] = {
                    "success": False,
                    "message": error_msg,
                    "stats": {},
                    "status": "error"
                }
                
                self.table_finished.emit(table_name, False, error_msg, {})
        
        # Resumen final
        total_records = sum(r["stats"].get("records_count", 0) for r in self.results.values() if r["success"])
        
        summary = {
            "total_tables": total_tables,
            "updated": updated_count,
            "skipped": skipped_count,
            "errors": error_count,
            "total_records": total_records,
            "results": self.results
        }
        
        self.progress_update.emit(
            f"🎉 Actualización masiva completada: "
            f"✅ {updated_count} actualizadas, ℹ️ {skipped_count} sin cambios, ❌ {error_count} errores"
        )
        
        self.all_finished.emit(summary)
    
    def update_single_table(self, table_name, config):
        """Actualiza una sola tabla (método interno)"""
        # Verificar archivo
        file_path = config["source_file"]
        if not os.path.exists(file_path):
            return False, "Archivo no encontrado", {}
        
        # Verificar cambios
        tracking_data = self.load_tracking()
        table_data = tracking_data.get(table_name, {})
        
        current_hash = self.get_file_hash(file_path)
        stored_hash = table_data.get("last_hash", "")
        
        current_modified = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
        
        has_changed = current_hash != stored_hash
        
        if not has_changed and not self.force_update:
            return False, "No se detectaron cambios en el archivo", {
                "last_updated": table_data.get("last_updated", "N/A"),
                "records_count": table_data.get("records_count", 0)
            }
        
        # Procesar archivo
        df = read_file_data(file_path, config)
        detected_format = detect_column_format(file_path, config)
        df_processed = process_dataframe_columns(df, config, detected_format)
        
        # Actualizar BD
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # Crear tabla
            cursor.execute(config["create_table_sql"])
            
            # Limpiar y insertar
            cursor.execute(f"DELETE FROM {config['table_name']}")
            df_processed.to_sql(config['table_name'], conn, if_exists='append', index=False)
            conn.commit()
            
            # Actualizar tracking
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            tracking_data[table_name] = {
                "file_path": file_path,
                "last_hash": current_hash,
                "last_modified": current_modified,
                "last_updated": now,
                "records_count": len(df_processed),
                "detected_format": detected_format
            }
            self.save_tracking(tracking_data)
            
            return True, f"Actualizada exitosamente", {
                "records_count": len(df_processed),
                "last_updated": now,
                "detected_format": detected_format
            }
            
        finally:
            conn.close()

class R4EasyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.available_tables = get_all_tables()
        self.current_table = self.available_tables[0] if self.available_tables else "expedite"
        self.updater_thread = None
        self.bulk_updater_thread = None
        
        self.setWindowTitle(f"R4 Easy - {len(self.available_tables)} Tablas")
        self.setGeometry(100, 100, 1000, 700)
        
        self.setup_ui()
        self.load_initial_data()
        
    def setup_ui(self):
        """Configura la interfaz con pestañas"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)
        
        # Título
        title_label = QLabel(f"🤖 R4 Easy - Auto Database Manager")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin: 10px; padding: 10px;")
        layout.addWidget(title_label)
        
        # Crear pestañas
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)
        
        # Pestaña 1: Control Individual
        self.setup_individual_tab()
        
        # Pestaña 2: Actualización Masiva
        self.setup_bulk_tab()
        
    def setup_individual_tab(self):
        """Configura la pestaña de control individual"""
        individual_widget = QWidget()
        layout = QVBoxLayout(individual_widget)
        
        # Selector de tabla
        table_group = QGroupBox("📋 Seleccionar Tabla")
        table_layout = QHBoxLayout(table_group)
        
        self.table_combo = QComboBox()
        self.table_combo.addItems(self.available_tables)
        self.table_combo.setCurrentText(self.current_table)
        self.table_combo.currentTextChanged.connect(self.on_table_changed)
        
        table_layout.addWidget(QLabel("Tabla activa:"))
        table_layout.addWidget(self.table_combo)
        table_layout.addStretch()
        
        layout.addWidget(table_group)
        
        # Estado actual
        status_group = QGroupBox("📈 Estado Actual")
        status_layout = QVBoxLayout(status_group)
        
        self.status_file_label = QLabel("📁 Archivo: --")
        self.status_modified_label = QLabel("📅 Última modificación: --")
        self.status_updated_label = QLabel("🕒 Última actualización BD: --")
        self.status_records_label = QLabel("📊 Registros en BD: --")
        
        status_layout.addWidget(self.status_file_label)
        status_layout.addWidget(self.status_modified_label)
        status_layout.addWidget(self.status_updated_label)
        status_layout.addWidget(self.status_records_label)
        
        layout.addWidget(status_group)
        
        # Controles individuales
        controls_group = QGroupBox("🎛️ Controles Individuales")
        controls_layout = QHBoxLayout(controls_group)
        
        self.update_btn = QPushButton("🔄 Verificar y Actualizar")
        self.update_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #2980b9; }
        """)
        self.update_btn.clicked.connect(self.start_update)
        
        self.force_update_btn = QPushButton("⚡ Forzar Actualización")
        self.force_update_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #c0392b; }
        """)
        self.force_update_btn.clicked.connect(self.force_update)
        
        self.refresh_btn = QPushButton("🔍 Actualizar Estado")
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #7f8c8d; }
        """)
        self.refresh_btn.clicked.connect(self.refresh_all_status)
        
        controls_layout.addWidget(self.update_btn)
        controls_layout.addWidget(self.force_update_btn)
        controls_layout.addWidget(self.refresh_btn)
        
        layout.addWidget(controls_group)
        
        # Progreso individual
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Log individual
        log_group = QGroupBox("📝 Log Individual")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #2c3e50;
                color: #ecf0f1;
                border: 1px solid #34495e;
                border-radius: 5px;
                padding: 10px;
                font-family: 'Courier New', monospace;
            }
        """)
        log_layout.addWidget(self.log_text)
        
        layout.addWidget(log_group)
        
        self.tab_widget.addTab(individual_widget, "🎯 Control Individual")
    
    def setup_bulk_tab(self):
        """Configura la pestaña de actualización masiva"""
        bulk_widget = QWidget()
        layout = QVBoxLayout(bulk_widget)
        
        # Información global
        info_group = QGroupBox(f"📊 Información Global - {len(self.available_tables)} Tablas Configuradas")
        info_layout = QVBoxLayout(info_group)
        
        info_text = f"Tablas disponibles: {', '.join(self.available_tables)}"
        info_label = QLabel(info_text)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        layout.addWidget(info_group)
        
        # Controles masivos
        bulk_controls_group = QGroupBox("🚀 Controles Masivos")
        bulk_controls_layout = QHBoxLayout(bulk_controls_group)
        
        self.bulk_update_btn = QPushButton("🔄 Actualizar Todas las Tablas")
        self.bulk_update_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                padding: 15px 25px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover { background-color: #229954; }
        """)
        self.bulk_update_btn.clicked.connect(self.start_bulk_update)
        
        self.force_bulk_update_btn = QPushButton("⚡ Forzar Actualización Completa")
        self.force_bulk_update_btn.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                border: none;
                padding: 15px 25px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover { background-color: #d35400; }
        """)
        self.force_bulk_update_btn.clicked.connect(self.force_bulk_update)
        
        bulk_controls_layout.addWidget(self.bulk_update_btn)
        bulk_controls_layout.addWidget(self.force_bulk_update_btn)
        
        layout.addWidget(bulk_controls_group)
        
        # Progreso masivo
        self.bulk_progress_bar = QProgressBar()
        self.bulk_progress_bar.setVisible(False)
        layout.addWidget(self.bulk_progress_bar)
        
        # Tabla de resultados
        results_group = QGroupBox("📋 Resultados por Tabla")
        results_layout = QVBoxLayout(results_group)
        
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(5)
        self.results_table.setHorizontalHeaderLabels([
            "Tabla", "Estado", "Registros", "Última Actualización", "Mensaje"
        ])
        self.results_table.horizontalHeader().setStretchLastSection(True)
        results_layout.addWidget(self.results_table)
        
        layout.addWidget(results_group)
        
        # Log masivo
        bulk_log_group = QGroupBox("📝 Log Masivo")
        bulk_log_layout = QVBoxLayout(bulk_log_group)
        
        self.bulk_log_text = QTextEdit()
        self.bulk_log_text.setMaximumHeight(150)
        self.bulk_log_text.setStyleSheet("""
            QTextEdit {
                background-color: #2c3e50;
                color: #ecf0f1;
                border: 1px solid #34495e;
                border-radius: 5px;
                padding: 10px;
                font-family: 'Courier New', monospace;
            }
        """)
        bulk_log_layout.addWidget(self.bulk_log_text)
        
        layout.addWidget(bulk_log_group)
        
        self.tab_widget.addTab(bulk_widget, "🚀 Actualización Masiva")
    
    def load_initial_data(self):
        """Carga datos iniciales"""
        self.refresh_status()
        self.refresh_results_table()
        self.log_message(f"✅ R4 Easy inicializado con {len(self.available_tables)} tablas")
    
    def log_message(self, message):
        """Agrega mensaje al log individual"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_text.append(formatted_message)
        
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def bulk_log_message(self, message):
        """Agrega mensaje al log masivo"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.bulk_log_text.append(formatted_message)
        
        scrollbar = self.bulk_log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def on_table_changed(self, table_name):
        """Cambia la tabla seleccionada"""
        self.current_table = table_name
        self.log_message(f"📋 Tabla cambiada a: {table_name}")
        self.refresh_status()
    
    def refresh_status(self):
        """Actualiza el estado de la tabla individual"""
        try:
            config = TABLES_CONFIG.get(self.current_table)
            if not config:
                return
            
            # Info del archivo
            source_file = config["source_file"]
            self.status_file_label.setText(f"📁 Archivo: {os.path.basename(source_file)}")
            
            if os.path.exists(source_file):
                modified_time = datetime.fromtimestamp(os.path.getmtime(source_file)).strftime("%Y-%m-%d %H:%M:%S")
                self.status_modified_label.setText(f"📅 Última modificación: {modified_time}")
            else:
                self.status_modified_label.setText("📅 Última modificación: Archivo no encontrado")
            
            # Info de la BD
            tracking_file = BASE_PATHS["tracking_file"]
            if os.path.exists(tracking_file):
                with open(tracking_file, 'r') as f:
                    tracking_data = json.load(f)
                
                table_data = tracking_data.get(self.current_table, {})
                last_updated = table_data.get("last_updated", "N/A")
                records_count = table_data.get("records_count", 0)
                
                self.status_updated_label.setText(f"🕒 Última actualización BD: {last_updated}")
                self.status_records_label.setText(f"📊 Registros en BD: {records_count}")
            else:
                self.status_updated_label.setText("🕒 Última actualización BD: N/A")
                self.status_records_label.setText("📊 Registros en BD: N/A")
                
        except Exception as e:
            self.log_message(f"❌ Error actualizando estado: {str(e)}")
    
    def refresh_all_status(self):
        """Actualiza todos los estados y muestra confirmación"""
        try:
            self.log_message("🔍 Actualizando estado del sistema...")
            
            # Actualizar estado individual
            self.refresh_status()
            
            # Actualizar tabla de resultados
            self.refresh_results_table()
            
            # Mostrar información actualizada
            now = datetime.now().strftime("%H:%M:%S")
            self.log_message(f"✅ Estado actualizado a las {now}")
            
            # Contar archivos disponibles
            available_files = 0
            total_records = 0
            
            tracking_file = BASE_PATHS["tracking_file"]
            if os.path.exists(tracking_file):
                with open(tracking_file, 'r') as f:
                    tracking_data = json.load(f)
                
                for table_name, config in TABLES_CONFIG.items():
                    if os.path.exists(config["source_file"]):
                        available_files += 1
                    
                    table_data = tracking_data.get(table_name, {})
                    total_records += table_data.get("records_count", 0)
            else:
                for table_name, config in TABLES_CONFIG.items():
                    if os.path.exists(config["source_file"]):
                        available_files += 1
            
            # Mostrar resumen en el log
            self.log_message(f"📊 Resumen: {available_files}/{len(TABLES_CONFIG)} archivos encontrados")
            self.log_message(f"📈 Total registros en BD: {total_records:,}")
            
            # Mostrar mensaje de confirmación
            QMessageBox.information(
                self,
                "Estado Actualizado", 
                f"🔍 Estado del sistema actualizado\n\n"
                f"📊 Archivos encontrados: {available_files}/{len(TABLES_CONFIG)}\n"
                f"📈 Registros totales: {total_records:,}\n"
                f"🕒 Actualizado: {now}"
            )
            
        except Exception as e:
            self.log_message(f"❌ Error actualizando estado: {str(e)}")
            QMessageBox.warning(self, "Error", f"Error actualizando estado:\n{str(e)}")
        try:
            config = TABLES_CONFIG.get(self.current_table)
            if not config:
                return
            
            # Info del archivo
            source_file = config["source_file"]
            self.status_file_label.setText(f"📁 Archivo: {os.path.basename(source_file)}")
            
            if os.path.exists(source_file):
                modified_time = datetime.fromtimestamp(os.path.getmtime(source_file)).strftime("%Y-%m-%d %H:%M:%S")
                self.status_modified_label.setText(f"📅 Última modificación: {modified_time}")
            else:
                self.status_modified_label.setText("📅 Última modificación: Archivo no encontrado")
            
            # Info de la BD
            tracking_file = BASE_PATHS["tracking_file"]
            if os.path.exists(tracking_file):
                with open(tracking_file, 'r') as f:
                    tracking_data = json.load(f)
                
                table_data = tracking_data.get(self.current_table, {})
                last_updated = table_data.get("last_updated", "N/A")
                records_count = table_data.get("records_count", 0)
                
                self.status_updated_label.setText(f"🕒 Última actualización BD: {last_updated}")
                self.status_records_label.setText(f"📊 Registros en BD: {records_count}")
            else:
                self.status_updated_label.setText("🕒 Última actualización BD: N/A")
                self.status_records_label.setText("📊 Registros en BD: N/A")
                
        except Exception as e:
            self.log_message(f"❌ Error actualizando estado: {str(e)}")
    
    def refresh_results_table(self):
        """Actualiza la tabla de resultados"""
        try:
            tracking_file = BASE_PATHS["tracking_file"]
            tracking_data = {}
            
            if os.path.exists(tracking_file):
                with open(tracking_file, 'r') as f:
                    tracking_data = json.load(f)
            
            self.results_table.setRowCount(len(self.available_tables))
            
            for i, table_name in enumerate(self.available_tables):
                config = TABLES_CONFIG[table_name]
                table_data = tracking_data.get(table_name, {})
                
                # Nombre de tabla
                self.results_table.setItem(i, 0, QTableWidgetItem(table_name))
                
                # Estado
                file_exists = os.path.exists(config["source_file"])
                last_updated = table_data.get("last_updated", "")
                
                if not file_exists:
                    status = "❌ Archivo no encontrado"
                elif last_updated:
                    status = "✅ Actualizada"
                else:
                    status = "⚠️ Pendiente"
                
                self.results_table.setItem(i, 1, QTableWidgetItem(status))
                
                # Registros
                records = table_data.get("records_count", 0)
                self.results_table.setItem(i, 2, QTableWidgetItem(f"{records:,}"))
                
                # Última actualización
                self.results_table.setItem(i, 3, QTableWidgetItem(last_updated or "N/A"))
                
                # Mensaje
                if not file_exists:
                    message = "Archivo no encontrado"
                elif last_updated:
                    message = f"OK - {table_data.get('detected_format', 'N/A')}"
                else:
                    message = "Sin procesar"
                
                self.results_table.setItem(i, 4, QTableWidgetItem(message))
                
        except Exception as e:
            self.bulk_log_message(f"❌ Error actualizando tabla de resultados: {str(e)}")
    
    def start_update(self):
        """Inicia actualización individual"""
        if self.updater_thread and self.updater_thread.isRunning():
            QMessageBox.warning(self, "Actualización en Proceso", "Ya hay una actualización en curso.")
            return
        
        self.log_message(f"🔄 Iniciando verificación de {self.current_table}...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        
        self.update_btn.setEnabled(False)
        self.force_update_btn.setEnabled(False)
        
        self.updater_thread = SimpleUpdater(self.current_table)
        self.updater_thread.progress_update.connect(self.log_message)
        self.updater_thread.finished_update.connect(self.on_update_finished)
        self.updater_thread.start()
    
    def force_update(self):
        """Fuerza actualización individual"""
        reply = QMessageBox.question(
            self,
            "Forzar Actualización",
            f"⚠️ ¿Forzar actualización completa de '{self.current_table}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                tracking_file = BASE_PATHS["tracking_file"]
                if os.path.exists(tracking_file):
                    with open(tracking_file, 'r') as f:
                        tracking_data = json.load(f)
                    
                    if self.current_table in tracking_data:
                        tracking_data[self.current_table]["last_hash"] = ""
                    
                    with open(tracking_file, 'w') as f:
                        json.dump(tracking_data, f, indent=2)
                
                self.start_update()
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error forzando actualización: {str(e)}")
    
    def start_bulk_update(self):
        """Inicia actualización masiva"""
        if self.bulk_updater_thread and self.bulk_updater_thread.isRunning():
            QMessageBox.warning(self, "Actualización en Proceso", "Ya hay una actualización masiva en curso.")
            return
        
        reply = QMessageBox.question(
            self,
            "Actualización Masiva",
            f"🚀 ¿Iniciar actualización de todas las {len(self.available_tables)} tablas?\n\n"
            "Solo se actualizarán las tablas con cambios detectados.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            self.bulk_log_message(f"🚀 Iniciando actualización masiva de {len(self.available_tables)} tablas...")
            self.bulk_progress_bar.setVisible(True)
            self.bulk_progress_bar.setRange(0, len(self.available_tables))
            self.bulk_progress_bar.setValue(0)
            
            self.bulk_update_btn.setEnabled(False)
            self.force_bulk_update_btn.setEnabled(False)
            
            self.bulk_updater_thread = BulkUpdater(force_update=False)
            self.bulk_updater_thread.progress_update.connect(self.bulk_log_message)
            self.bulk_updater_thread.table_finished.connect(self.on_table_finished)
            self.bulk_updater_thread.all_finished.connect(self.on_bulk_finished)
            self.bulk_updater_thread.start()
    
    def force_bulk_update(self):
        """Fuerza actualización masiva completa"""
        reply = QMessageBox.question(
            self,
            "Forzar Actualización Completa",
            f"⚡ ¿Forzar actualización completa de TODAS las {len(self.available_tables)} tablas?\n\n"
            "⚠️ Esto recreará todas las tablas sin verificar cambios.\n"
            "Puede tomar varios minutos.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.bulk_log_message(f"⚡ Forzando actualización completa de {len(self.available_tables)} tablas...")
            self.bulk_progress_bar.setVisible(True)
            self.bulk_progress_bar.setRange(0, len(self.available_tables))
            self.bulk_progress_bar.setValue(0)
            
            self.bulk_update_btn.setEnabled(False)
            self.force_bulk_update_btn.setEnabled(False)
            
            self.bulk_updater_thread = BulkUpdater(force_update=True)
            self.bulk_updater_thread.progress_update.connect(self.bulk_log_message)
            self.bulk_updater_thread.table_finished.connect(self.on_table_finished)
            self.bulk_updater_thread.all_finished.connect(self.on_bulk_finished)
            self.bulk_updater_thread.start()
    
    def on_update_finished(self, success, message, stats):
        """Maneja finalización de actualización individual"""
        self.progress_bar.setVisible(False)
        self.update_btn.setEnabled(True)
        self.force_update_btn.setEnabled(True)
        
        if success:
            self.log_message(f"✅ {message}")
            QMessageBox.information(
                self,
                "Actualización Completada",
                f"✅ {message}\n\n📊 Registros: {stats.get('records_count', 'N/A')}"
            )
        else:
            self.log_message(f"ℹ️ {message}")
            if "No se detectaron cambios" in message:
                QMessageBox.information(self, "Sin Cambios", f"ℹ️ {message}")
            else:
                QMessageBox.warning(self, "Error", f"❌ {message}")
        
        self.refresh_status()
        self.refresh_results_table()
    
    def on_table_finished(self, table_name, success, message, stats):
        """Maneja finalización de una tabla en actualización masiva"""
        current_value = self.bulk_progress_bar.value()
        self.bulk_progress_bar.setValue(current_value + 1)
        
        # Actualizar tabla de resultados en tiempo real
        self.refresh_results_table()
    
    def on_bulk_finished(self, summary):
        """Maneja finalización de actualización masiva"""
        self.bulk_progress_bar.setVisible(False)
        self.bulk_update_btn.setEnabled(True)
        self.force_bulk_update_btn.setEnabled(True)
        
        if "error" in summary:
            QMessageBox.critical(self, "Error Crítico", f"❌ Error durante actualización masiva:\n{summary['error']}")
            return
        
        # Mostrar resumen
        total = summary["total_tables"]
        updated = summary["updated"]
        skipped = summary["skipped"]
        errors = summary["errors"]
        total_records = summary["total_records"]
        
        self.bulk_log_message(
            f"🎉 Resumen final: {updated} actualizadas, {skipped} sin cambios, {errors} errores"
        )
        
        # Actualizar tabla de resultados
        self.refresh_results_table()
        
        # Mostrar diálogo de resumen
        if errors == 0:
            icon = QMessageBox.Information
            title = "Actualización Masiva Completada"
            if updated > 0:
                message = (
                    f"🎉 ¡Actualización masiva exitosa!\n\n"
                    f"✅ Tablas actualizadas: {updated}\n"
                    f"ℹ️ Sin cambios: {skipped}\n"
                    f"📊 Total de registros procesados: {total_records:,}"
                )
            else:
                message = (
                    f"ℹ️ Actualización masiva completada\n\n"
                    f"Todas las tablas ya estaban actualizadas.\n"
                    f"📊 Total de registros en BD: {total_records:,}"
                )
        else:
            icon = QMessageBox.Warning
            title = "Actualización Masiva con Errores"
            message = (
                f"⚠️ Actualización completada con algunos errores\n\n"
                f"✅ Actualizadas: {updated}\n"
                f"ℹ️ Sin cambios: {skipped}\n"
                f"❌ Con errores: {errors}\n"
                f"📊 Registros procesados: {total_records:,}\n\n"
                f"Revisa el log para más detalles."
            )
        
        msg_box = QMessageBox(icon, title, message, QMessageBox.Ok, self)
        msg_box.exec_()
        
        # Actualizar estado individual si está en la tabla actual
        self.refresh_status()

def main():
    """Función principal"""
    app = QApplication(sys.argv)
    
    if not TABLES_CONFIG:
        QMessageBox.critical(
            None, 
            "Error de Configuración", 
            "No se encontraron tablas en config_tables.py"
        )
        sys.exit(1)
    
    app.setStyleSheet("""
        QMainWindow { background-color: #ecf0f1; }
        QGroupBox {
            font-weight: bold;
            border: 2px solid #bdc3c7;
            border-radius: 10px;
            margin-top: 1ex;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        QLabel { color: #2c3e50; }
        QTabWidget::pane {
            border: 1px solid #bdc3c7;
            border-radius: 5px;
        }
        QTabBar::tab {
            background-color: #95a5a6;
            color: white;
            padding: 8px 16px;
            margin: 2px;
            border-radius: 5px;
        }
        QTabBar::tab:selected {
            background-color: #3498db;
        }
        QTabBar::tab:hover {
            background-color: #7f8c8d;
        }
        QTableWidget {
            gridline-color: #bdc3c7;
            background-color: white;
            alternate-background-color: #f8f9fa;
        }
        QTableWidget::item {
            padding: 5px;
        }
    """)
    
    try:
        window = R4EasyApp()
        window.show()
        
        # Mostrar mensaje de bienvenida
        QMessageBox.information(
            window,
            "R4 Easy - Sistema Inicializado",
            f"🤖 R4 Easy iniciado exitosamente!\n\n"
            f"📊 Tablas configuradas: {len(TABLES_CONFIG)}\n"
            f"📋 Disponibles: {', '.join(TABLES_CONFIG.keys())}\n\n"
            f"💡 Usa las pestañas para:\n"
            f"   • Control Individual: Actualizar tabla por tabla\n"
            f"   • Actualización Masiva: Procesar todas de una vez"
        )
        
        sys.exit(app.exec_())
        
    except Exception as e:
        QMessageBox.critical(None, "Error Fatal", f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()