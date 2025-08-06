#!/usr/bin/env python3
"""
Funciones de backup para R4 Easy - Sistema Simple de Base de Datos
Agrega funcionalidad de copia automática al escritorio después de actualizaciones masivas
"""

import shutil
import os
import sys
import platform
import subprocess
from datetime import datetime
from PyQt5.QtWidgets import QMessageBox

def get_desktop_path():
    """Obtiene la ruta del escritorio del usuario de manera multiplataforma"""
    try:
        # Windows
        if platform.system() == 'Windows':
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        # macOS y Linux
        else:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        
        # Verificar si existe la carpeta Desktop, si no usar el home
        if not os.path.exists(desktop):
            desktop = os.path.expanduser("~")
        
        return desktop
    except Exception:
        # Fallback al directorio home si hay problemas
        return os.path.expanduser("~")

def create_database_backup(db_path, custom_path=None):
    """
    Crea una copia de la base de datos con timestamp
    
    Args:
        db_path (str): Ruta de la base de datos original
        custom_path (str, optional): Ruta personalizada para el backup. 
                                Si no se especifica, usa el escritorio.
    
    Returns:
        tuple: (success: bool, backup_path: str, message: str)
    """
    try:
        if not os.path.exists(db_path):
            return False, "", "La base de datos original no existe"
        
        # Determinar destino
        if custom_path:
            backup_dir = custom_path
        else:
            backup_dir = get_desktop_path()
        
        # Crear nombre del backup con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        db_filename = os.path.basename(db_path)
        name_without_ext = os.path.splitext(db_filename)[0]
        ext = os.path.splitext(db_filename)[1]
        
        backup_filename = f"{name_without_ext}{ext}"
        backup_path = os.path.join(backup_dir, backup_filename)
        
        # Asegurar que el directorio destino existe
        os.makedirs(backup_dir, exist_ok=True)
        
        # Copiar archivo
        shutil.copy2(db_path, backup_path)
        
        # Verificar que la copia fue exitosa
        if os.path.exists(backup_path):
            original_size = os.path.getsize(db_path)
            backup_size = os.path.getsize(backup_path)
            
            if original_size == backup_size:
                return True, backup_path, f"Backup creado exitosamente en: {backup_path}"
            else:
                return False, backup_path, "Error: El tamaño del backup no coincide con el original"
        else:
            return False, backup_path, "Error: No se pudo crear el archivo de backup"
            
    except PermissionError:
        return False, "", f"Error de permisos: No se puede escribir en {backup_dir}"
    except Exception as e:
        return False, "", f"Error inesperado creando backup: {str(e)}"

def show_backup_dialog(parent, success, backup_path, message, stats=None):
    """
    Muestra diálogo informativo sobre el backup
    
    Args:
        parent: Widget padre para el diálogo
        success (bool): Si el backup fue exitoso
        backup_path (str): Ruta del archivo de backup
        message (str): Mensaje descriptivo
        stats (dict, optional): Estadísticas de la actualización
    """
    try:
        if success:
            # Obtener información del archivo
            file_size = os.path.getsize(backup_path) / (1024 * 1024)  # MB
            backup_time = datetime.now().strftime("%H:%M:%S")
            
            dialog_message = (
                f"💾 ¡Backup creado exitosamente!\n\n"
                f"📁 Ubicación: {backup_path}\n"
                f"📊 Tamaño: {file_size:.2f} MB\n"
                f"🕒 Hora: {backup_time}\n"
            )
            
            if stats:
                dialog_message += (
                    f"\n📈 Estadísticas de actualización:\n"
                    f"   • Tablas actualizadas: {stats.get('updated', 0)}\n"
                    f"   • Registros totales: {stats.get('total_records', 0):,}\n"
                )
            
            dialog_message += f"\n✅ {message}"
            
            msg_box = QMessageBox(QMessageBox.Information, "Backup Completado", dialog_message, QMessageBox.Ok, parent)
            
            # Agregar botón para abrir la carpeta del backup
            open_folder_btn = msg_box.addButton("📁 Abrir Carpeta", QMessageBox.ActionRole)
            
            msg_box.exec_()
            
            # Si el usuario clickeó en abrir carpeta
            if msg_box.clickedButton() == open_folder_btn:
                open_backup_folder(backup_path)
                
        else:
            QMessageBox.warning(
                parent,
                "Error en Backup",
                f"❌ No se pudo crear el backup\n\n{message}"
            )
            
    except Exception as e:
        QMessageBox.critical(
            parent,
            "Error",
            f"Error mostrando diálogo de backup: {str(e)}"
        )

def open_backup_folder(backup_path):
    """Abre la carpeta que contiene el backup"""
    try:
        backup_dir = os.path.dirname(backup_path)
        
        # Abrir carpeta según el sistema operativo
        if platform.system() == 'Windows':
            os.startfile(backup_dir)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.run(['open', backup_dir])
        else:  # Linux y otros Unix
            subprocess.run(['xdg-open', backup_dir])
                
    except Exception as e:
        print(f"No se pudo abrir la carpeta: {str(e)}")

# Función para integrar en BulkUpdater
def create_post_update_backup(db_path, stats, parent_widget=None):
    """
    Función específica para crear backup después de actualización masiva
    
    Args:
        db_path (str): Ruta de la base de datos
        stats (dict): Estadísticas de la actualización
        parent_widget: Widget padre para diálogos
    
    Returns:
        tuple: (success: bool, backup_path: str, message: str)
    """
    # Solo crear backup si hubo actualizaciones
    if stats.get('updated', 0) > 0:
        success, backup_path, message = create_database_backup(db_path)
        
        if parent_widget:
            show_backup_dialog(parent_widget, success, backup_path, message, stats)
        
        return success, backup_path, message
    else:
        message = "No se creó backup - no hubo actualizaciones"
        if parent_widget:
            # Opcional: mostrar mensaje informativo
            QMessageBox.information(
                parent_widget,
                "Backup No Necesario",
                f"ℹ️ No se creó backup automático\n\n"
                f"Razón: No se detectaron cambios en ninguna tabla.\n"
                f"📊 Total registros en BD: {stats.get('total_records', 0):,}"
            )
        
        return False, "", message