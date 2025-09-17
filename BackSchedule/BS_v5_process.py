import sqlite3
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from datetime import datetime
import threading
import time
import traceback
import re

class BOMExcelProcessor:
    """Procesador de BOM con lógica completa integrada"""
    
    def __init__(self):
        self.df_bom = None
        self.df_resultado = None
        
    def extract_level_safe(self, level_str):
        """Extrae el número del nivel de forma segura"""
        if pd.isna(level_str):
            return 0
        
        # Si ya es un número, devolverlo
        if isinstance(level_str, (int, float)):
            return int(level_str)
        
        # Si es string, buscar el número
        match = re.search(r'(\d+)', str(level_str))
        if match:
            return int(match.group(1))
        return 0
    
    def crear_lt_final_parchado(self, df_trabajo):
        """
        Crea la columna LT_Final_parchado antes del procesamiento del árbol
        Lógica: Si LT_No_in_R4 ≠ 0 usar LT_No_in_R4, si no usar LT_in92
        """
        # PASO 0: Llenar NULLs con 0 en LT_No_in_R4 antes de procesar
        df_trabajo['LT_No_in_R4'] = df_trabajo['LT_No_in_R4'].fillna(0)
        
        def calcular_lt_parchado(row):
            lt_no_in_r4 = row['LT_No_in_R4']
            lt_in92 = row['LT_in92']
            
            # Convertir a float de forma segura
            try:
                lt_no_in_r4_val = float(lt_no_in_r4) if pd.notna(lt_no_in_r4) and str(lt_no_in_r4).strip() != '' else 0.0
            except:
                lt_no_in_r4_val = 0.0
            
            try:
                lt_in92_val = float(lt_in92) if pd.notna(lt_in92) else 0.0
            except:
                lt_in92_val = 0.0
            
            # Lógica: Si LT_No_in_R4 ≠ 0, usar ese valor, si no usar LT_in92
            if lt_no_in_r4_val != 0:
                return lt_no_in_r4_val
            else:
                return lt_in92_val
        
        df_trabajo['LT_Final_parchado'] = df_trabajo.apply(calcular_lt_parchado, axis=1)
        return df_trabajo
    
    def procesar_bom_completo(self, df_bom):
        """Procesa el BOM con toda la lógica completa"""
        df_trabajo = df_bom.copy()
        
        # PASO 1: Crear LT_Final_parchado
        df_trabajo = self.crear_lt_final_parchado(df_trabajo)
        
        # Limpiar Level_Number si tiene formato con puntos
        df_trabajo['Level_Number_Clean'] = df_trabajo['Level_Number'].apply(self.extract_level_safe)
        
        # Crear columnas nuevas
        df_trabajo['arbol'] = ''
        df_trabajo['arbol_Plantype'] = ''
        df_trabajo['arbol_LT'] = ''
        df_trabajo['LT_sistema'] = 0.0
        df_trabajo['LT_Release'] = 0.0
        
        # PASO 2: PROCESAR HERENCIA DE LT_PlanType PRIMERO
        lt_plantype_heredado = None
        nivel_lt_plantype_heredado = 0
        
        for idx, row in df_trabajo.iterrows():
            nivel_actual = df_trabajo.at[idx, 'Level_Number_Clean']
            if nivel_actual == 0:
                continue
                
            lt_plantype_original = float(row['LT_PlanType']) if pd.notna(row['LT_PlanType']) else 0.0
            
            # Limpiar herencia si volvemos a nivel igual o superior
            if lt_plantype_heredado is not None and nivel_actual <= nivel_lt_plantype_heredado:
                lt_plantype_heredado = None
                nivel_lt_plantype_heredado = 0
            
            # Si tiene LT_PlanType, establecer herencia
            if lt_plantype_original > 0:
                lt_plantype_heredado = lt_plantype_original
                nivel_lt_plantype_heredado = nivel_actual
                df_trabajo.at[idx, 'LT_PlanType'] = lt_plantype_original
            # Si hay herencia activa
            elif lt_plantype_heredado is not None:
                df_trabajo.at[idx, 'LT_PlanType'] = lt_plantype_heredado
            else:
                df_trabajo.at[idx, 'LT_PlanType'] = lt_plantype_original
        
        # PASO 3: PROCESAR ÁRBOL Y CALCULAR LT_sistema y LT_Release
        # Diccionarios para almacenar por nivel
        nivel_componente = {}
        nivel_plantype = {}
        nivel_lt_final_parchado = {}
        
        # Iterar por cada fila para calcular LTs
        for idx, row in df_trabajo.iterrows():
            nivel_actual = row['Level_Number_Clean']
            
            if nivel_actual == 0:
                continue
            
            componente_actual = row['Component']
            plantype_actual = row['Plan_Type']
            lt_final_parchado_actual = row['LT_Final_parchado']
            mli_valor = str(row['MLI']).upper().strip() if pd.notna(row['MLI']) else ""
            
            # Actualizar diccionarios
            nivel_componente[nivel_actual] = componente_actual
            nivel_plantype[nivel_actual] = plantype_actual
            nivel_lt_final_parchado[nivel_actual] = lt_final_parchado_actual
            
            # Limpiar niveles superiores
            niveles_a_eliminar = [n for n in list(nivel_componente.keys()) if n > nivel_actual]
            for n in niveles_a_eliminar:
                if n in nivel_componente:
                    del nivel_componente[n]
                if n in nivel_plantype:
                    del nivel_plantype[n]
                if n in nivel_lt_final_parchado:
                    del nivel_lt_final_parchado[n]
            
            # Construir árbol
            if nivel_actual == 1:
                df_trabajo.at[idx, 'arbol'] = f"1...{componente_actual}"
                df_trabajo.at[idx, 'arbol_Plantype'] = f"1...{plantype_actual}" if pd.notna(plantype_actual) else ''
                df_trabajo.at[idx, 'LT_sistema'] = 0.0
                df_trabajo.at[idx, 'arbol_LT'] = "1...0"
                
                # LÓGICA PARA LT_Release EN NIVEL 1
                lt_plantype_procesado = df_trabajo.at[idx, 'LT_PlanType']
                
                if lt_plantype_procesado > 0:
                    lt_release_calculado = lt_plantype_procesado + 3
                    df_trabajo.at[idx, 'LT_Release'] = lt_release_calculado
                else:
                    if mli_valor == "L":
                        lt_release_calculado = 3
                        df_trabajo.at[idx, 'LT_Release'] = lt_release_calculado
                    else:
                        lt_release_calculado = lt_final_parchado_actual + 3
                        df_trabajo.at[idx, 'LT_Release'] = lt_release_calculado
            else:
                # Construir árbol de componentes y arbol_LT
                componentes_arbol = []
                plantypes_arbol = []
                lt_arbol = []
                suma_lt_sistema = 0.0
                running_sum = 0.0
                
                for nivel in sorted(nivel_lt_final_parchado.keys()):
                    if nivel < nivel_actual:
                        comp = nivel_componente[nivel]
                        pt = nivel_plantype[nivel]
                        componentes_arbol.append(f"{nivel}...{comp}")
                        if pd.notna(pt):
                            plantypes_arbol.append(f"{nivel}...{pt}")
                        
                        # Acumular LT
                        running_sum += nivel_lt_final_parchado[nivel]
                        suma_lt_sistema += nivel_lt_final_parchado[nivel]
                        lt_arbol.append(f"{nivel}...{running_sum}")
                
                df_trabajo.at[idx, 'arbol'] = ', '.join(componentes_arbol)
                df_trabajo.at[idx, 'arbol_Plantype'] = ', '.join(plantypes_arbol)
                df_trabajo.at[idx, 'arbol_LT'] = ', '.join(lt_arbol)
                df_trabajo.at[idx, 'LT_sistema'] = suma_lt_sistema
                
                # LÓGICA PARA LT_Release CON MLI
                lt_plantype_procesado = df_trabajo.at[idx, 'LT_PlanType']
                
                if lt_plantype_procesado > 0:
                    lt_release_calculado = lt_plantype_procesado + 3
                    df_trabajo.at[idx, 'LT_Release'] = lt_release_calculado
                else:
                    if mli_valor == "L":
                        lt_release_calculado = suma_lt_sistema + 3
                        df_trabajo.at[idx, 'LT_Release'] = lt_release_calculado
                    else:
                        lt_release_calculado = suma_lt_sistema + lt_final_parchado_actual + 3
                        df_trabajo.at[idx, 'LT_Release'] = lt_release_calculado
        
        # Eliminar columna temporal
        df_trabajo = df_trabajo.drop('Level_Number_Clean', axis=1)
        return df_trabajo

class BOMTreeBuilder:
    def __init__(self):
        self.db_path = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\R4Database\R4Database.db"
        self.df_bom = None
        self.df_resultado = None
        self.processor = BOMExcelProcessor()

    def cargar_datos_bom(self):
        """Carga los datos del BOM desde la base de datos y añade LT desde in92"""
        try:
            conn = sqlite3.connect(self.db_path)
            query = """
            SELECT b.id,
                    b.Level_Number,
                    b.Component,
                    b.Description,
                    b.CE,
                    b.T,
                    b.Sort,
                    b.UM,
                    b.MLI,
                    b.Plan_Type,
                    b.Unit_Qty,
                    b.Std_Cost,
                    b.Ext_Std,
                    b.Labor_IN,
                    b.Lab_Rem,
                    b.Mat_IN,
                    b.Mat_Rem,
                    b.This_lvl,
                    b.Lab_Hrs,
                    b.key,
                    b.Orden_BOM_Original,
                    i.LT AS LT_in92,
                    p.LT_PlanType AS LT_PlanType,
                    l.LeadTime AS LT_No_in_R4
                FROM bom b
                LEFT JOIN in92 i ON b.Component = i.ItemNo
                LEFT JOIN lt_by_process p ON b.Plan_Type = p.PlanType
                LEFT JOIN leadtimes_parchados l ON b.Component = l.ItemNo;
            """
            self.df_bom = pd.read_sql_query(query, conn)
            conn.close()
            
            # Ordenar por Orden_BOM_Original
            self.df_bom = self.df_bom.sort_values('Orden_BOM_Original').reset_index(drop=True)
            return True
            
        except Exception as e:
            raise Exception(f"Error al cargar la base de datos: {str(e)}")
    
    def obtener_keys_unicos(self):
        """Obtiene todos los keys únicos ordenados"""
        if self.df_bom is None:
            return []
        
        # Obtener keys únicos manteniendo el orden de aparición
        keys_ordenados = self.df_bom['key'].drop_duplicates().tolist()
        return [k for k in keys_ordenados if pd.notna(k)]
    
    def obtener_bom_por_key(self, key):
        """Obtiene todos los registros de un key específico"""
        # Filtrar por key
        df_key = self.df_bom[self.df_bom['key'] == key].copy()
        
        if df_key.empty:
            return None
        
        # Obtener el primer índice y nivel
        idx_inicial = df_key.index[0]
        nivel_inicial = df_key.iloc[0]['Level_Number']
        
        # Lista para almacenar los índices a incluir
        indices_incluir = list(df_key.index)
        
        # Buscar todos los hijos del key
        for idx in range(idx_inicial + 1, len(self.df_bom)):
            if idx in indices_incluir:
                continue
            if self.df_bom.iloc[idx]['Level_Number'] > nivel_inicial:
                indices_incluir.append(idx)
            else:
                # Si encontramos un item del mismo nivel o superior, terminamos
                break
        
        # Filtrar el DataFrame con todos los índices
        df_filtrado = self.df_bom.loc[indices_incluir].copy()
        df_filtrado = df_filtrado.sort_values('Orden_BOM_Original').reset_index(drop=True)
        
        return df_filtrado
    
    def crear_arbol_bom(self, df_filtrado):
        """Crea el árbol jerárquico del BOM con lógica completa"""
        return self.processor.procesar_bom_completo(df_filtrado)

class ProgressWindow:
    def __init__(self, parent, total_keys):
        self.parent = parent
        self.total_keys = total_keys
        self.current_key_index = 0
        self.start_time = time.time()
        self.paused = False
        self.cancelled = False
        self.error_occurred = False
        self.continue_on_error = False
        
        # Crear ventana de progreso
        self.window = tk.Toplevel(parent)
        self.window.title("Procesando BOM")
        self.window.geometry("600x500")
        self.window.resizable(False, False)
        
        # Centrar ventana
        self.center_window()
        
        # Prevenir cierre accidental
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.create_widgets()
        
        # Log de errores
        self.error_log = []
        
    def center_window(self):
        """Centra la ventana en la pantalla"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        """Crea los widgets de la ventana de progreso"""
        # Frame principal
        main_frame = ttk.Frame(self.window, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        self.lbl_titulo = ttk.Label(main_frame, text="Procesando BOM Completo", 
                                   font=('Arial', 14, 'bold'))
        self.lbl_titulo.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # Información del progreso
        self.lbl_key_actual = ttk.Label(main_frame, text="Preparando...", 
                                        font=('Arial', 10))
        self.lbl_key_actual.grid(row=1, column=0, columnspan=3, pady=5)
        
        # Contador de keys
        self.lbl_contador = ttk.Label(main_frame, text=f"Key 0 de {self.total_keys}")
        self.lbl_contador.grid(row=2, column=0, columnspan=3, pady=5)
        
        # Barra de progreso principal
        self.progress_bar = ttk.Progressbar(main_frame, length=550, mode='determinate')
        self.progress_bar.grid(row=3, column=0, columnspan=3, pady=10)
        self.progress_bar['maximum'] = self.total_keys
        
        # Porcentaje
        self.lbl_porcentaje = ttk.Label(main_frame, text="0%")
        self.lbl_porcentaje.grid(row=4, column=0, columnspan=3)
        
        # Tiempo
        self.lbl_tiempo = ttk.Label(main_frame, text="Tiempo transcurrido: 00:00:00 | Restante: Calculando...")
        self.lbl_tiempo.grid(row=5, column=0, columnspan=3, pady=10)
        
        # Frame de botones de control
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=6, column=0, columnspan=3, pady=15)
        
        # Botón Pausar/Reanudar
        self.btn_pausar = ttk.Button(control_frame, text="Pausar", 
                                     command=self.toggle_pause, width=15)
        self.btn_pausar.grid(row=0, column=0, padx=5)
        
        # Botón Cancelar
        self.btn_cancelar = ttk.Button(control_frame, text="Cancelar", 
                                       command=self.cancel_process, width=15)
        self.btn_cancelar.grid(row=0, column=1, padx=5)
        
        # Separador
        ttk.Separator(main_frame, orient='horizontal').grid(row=7, column=0, 
                                                           columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Frame para el log
        log_frame = ttk.LabelFrame(main_frame, text="Log de Proceso", padding="5")
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Text widget para el log con scrollbar
        self.log_text = tk.Text(log_frame, height=10, width=65, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar tags para colores
        self.log_text.tag_config("info", foreground="black")
        self.log_text.tag_config("success", foreground="green")
        self.log_text.tag_config("error", foreground="red")
        self.log_text.tag_config("warning", foreground="orange")
        
        # Frame de estadísticas
        stats_frame = ttk.LabelFrame(main_frame, text="Estadísticas", padding="5")
        stats_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.lbl_stats = ttk.Label(stats_frame, text="Procesados: 0 | Errores: 0 | Omitidos: 0")
        self.lbl_stats.grid(row=0, column=0)
        
    def update_progress(self, key_index, key_name, processed, errors, skipped):
        """Actualiza el progreso en la ventana"""
        if self.window.winfo_exists():
            self.current_key_index = key_index
            
            # Actualizar labels
            self.lbl_key_actual.config(text=f"Procesando: {key_name}")
            self.lbl_contador.config(text=f"Key {key_index + 1} de {self.total_keys}")
            
            # Actualizar barra de progreso
            self.progress_bar['value'] = key_index + 1
            
            # Calcular porcentaje
            porcentaje = ((key_index + 1) / self.total_keys) * 100
            self.lbl_porcentaje.config(text=f"{porcentaje:.1f}%")
            
            # Calcular tiempos
            tiempo_transcurrido = time.time() - self.start_time
            if key_index > 0:
                tiempo_por_key = tiempo_transcurrido / (key_index + 1)
                tiempo_restante = tiempo_por_key * (self.total_keys - key_index - 1)
            else:
                tiempo_restante = 0
            
            self.lbl_tiempo.config(text=f"Tiempo transcurrido: {self.format_time(tiempo_transcurrido)} | "
                                      f"Restante: {self.format_time(tiempo_restante)}")
            
            # Actualizar estadísticas
            self.lbl_stats.config(text=f"Procesados: {processed} | Errores: {errors} | Omitidos: {skipped}")
            
            # Actualizar la ventana
            self.window.update()
    
    def add_log(self, message, tag="info"):
        """Añade un mensaje al log"""
        if self.window.winfo_exists():
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", tag)
            self.log_text.see(tk.END)
            self.window.update()
    
    def format_time(self, seconds):
        """Formatea segundos a HH:MM:SS"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    def toggle_pause(self):
        """Pausa o reanuda el proceso"""
        self.paused = not self.paused
        if self.paused:
            self.btn_pausar.config(text="Reanudar")
            self.add_log("Proceso pausado por el usuario", "warning")
        else:
            self.btn_pausar.config(text="Pausar")
            self.add_log("Proceso reanudado", "info")
    
    def cancel_process(self):
        """Cancela el proceso"""
        if messagebox.askyesno("Confirmar", "¿Está seguro de cancelar el proceso?"):
            self.cancelled = True
            self.add_log("Proceso cancelado por el usuario", "error")
    
    def show_error(self, error_msg, key_name):
        """Muestra un error y pregunta qué hacer"""
        self.error_occurred = True
        self.add_log(f"ERROR en {key_name}: {error_msg}", "error")
        
        # Crear ventana de error
        error_window = tk.Toplevel(self.window)
        error_window.title("Error Encontrado")
        error_window.geometry("500x300")
        error_window.transient(self.window)
        error_window.grab_set()
        
        # Centrar ventana de error
        error_window.update_idletasks()
        width = error_window.winfo_width()
        height = error_window.winfo_height()
        x = (error_window.winfo_screenwidth() // 2) - (width // 2)
        y = (error_window.winfo_screenheight() // 2) - (height // 2)
        error_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Frame principal
        frame = ttk.Frame(error_window, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Mensaje de error
        ttk.Label(frame, text="Se encontró un error:", font=('Arial', 11, 'bold')).pack(pady=(0, 10))
        
        error_text = tk.Text(frame, height=8, width=55, wrap=tk.WORD)
        error_text.pack(pady=5)
        error_text.insert("1.0", f"Key: {key_name}\n\nError: {error_msg}")
        error_text.config(state='disabled')
        
        # Variable para la respuesta
        response = tk.StringVar(value="continue")
        
        def set_response(value):
            response.set(value)
            error_window.destroy()
        
        # Frame de botones
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=15)
        
        ttk.Button(btn_frame, text="Continuar con siguiente", 
                  command=lambda: set_response("continue")).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Reintentar este Key", 
                  command=lambda: set_response("retry")).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Cancelar todo", 
                  command=lambda: set_response("cancel")).grid(row=0, column=2, padx=5)
        
        # Checkbox para continuar automáticamente con errores
        self.continue_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Continuar automáticamente con futuros errores", 
                       variable=self.continue_var).pack(pady=5)
        
        # Esperar respuesta
        error_window.wait_window()
        
        self.continue_on_error = self.continue_var.get()
        
        return response.get()
    
    def on_closing(self):
        """Maneja el intento de cerrar la ventana"""
        if messagebox.askyesno("Confirmar", "¿Está seguro de cerrar? Se cancelará el proceso."):
            self.cancelled = True
            self.window.destroy()
    
    def finish(self, total_processed, total_errors):
        """Finaliza el proceso y muestra resumen"""
        self.add_log(f"Proceso completado: {total_processed} procesados, {total_errors} errores", 
                    "success" if total_errors == 0 else "warning")
        
        # Cambiar botones
        self.btn_pausar.config(state='disabled')
        self.btn_cancelar.config(text="Cerrar", command=self.window.destroy)
        
        # Mostrar mensaje final
        if total_errors == 0:
            messagebox.showinfo("Proceso Completado", 
                              f"El proceso se completó exitosamente.\n"
                              f"Total de keys procesados: {total_processed}")
        else:
            messagebox.showwarning("Proceso Completado con Errores", 
                                 f"El proceso se completó con algunos errores.\n"
                                 f"Keys procesados: {total_processed}\n"
                                 f"Errores encontrados: {total_errors}\n"
                                 f"Revise el log para más detalles.")

class BOMProcessor:
    def __init__(self, parent):
        self.parent = parent
        self.bom_builder = BOMTreeBuilder()
        self.progress_window = None
        self.db_nueva_path = None
        self.log_file_path = None
        self.info_file_path = None
    
    def validar_archivos_dependencias(self):
        """Valida los archivos de dependencias y retorna información detallada"""
        archivos_validar = [
            r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\SUPERBOM\SUPERBOM.txt",
            r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\Parche_Expedite\LeadTime_Parche.xlsx",
            r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\Parche_Expedite\PlanTypes_areas_primarias.xlsx"
        ]
        
        fecha_actual = datetime.now()
        resultados_validacion = []
        
        for archivo_path in archivos_validar:
            resultado = {
                'archivo': os.path.basename(archivo_path),
                'ruta_completa': archivo_path,
                'existe': False,
                'fecha_modificacion': 'No disponible',
                'peso_mb': 'No disponible',
                'dias_diferencia': 'No disponible',
                'estado': 'ERROR'
            }
            
            try:
                if os.path.exists(archivo_path):
                    resultado['existe'] = True
                    
                    # Obtener información del archivo
                    stat_archivo = os.stat(archivo_path)
                    
                    # Fecha de última modificación
                    fecha_modificacion = datetime.fromtimestamp(stat_archivo.st_mtime)
                    resultado['fecha_modificacion'] = fecha_modificacion.strftime('%Y-%m-%d %H:%M:%S')
                    
                    # Peso en MB
                    peso_bytes = stat_archivo.st_size
                    peso_mb = peso_bytes / (1024 * 1024)
                    resultado['peso_mb'] = f"{peso_mb:.2f} MB"
                    
                    # Diferencia en días
                    diferencia = fecha_actual - fecha_modificacion
                    resultado['dias_diferencia'] = diferencia.days
                    
                    # Determinar estado basado en antigüedad
                    if diferencia.days == 0:
                        resultado['estado'] = 'ACTUAL (mismo día)'
                    elif diferencia.days <= 7:
                        resultado['estado'] = 'RECIENTE (menos de 7 días)'
                    elif diferencia.days <= 30:
                        resultado['estado'] = 'ACEPTABLE (menos de 30 días)'
                    else:
                        resultado['estado'] = f'ANTIGUO ({diferencia.days} días)'
                        
                else:
                    resultado['estado'] = 'ARCHIVO NO ENCONTRADO'
                    
            except Exception as e:
                resultado['estado'] = f'ERROR: {str(e)}'
            
            resultados_validacion.append(resultado)
        
        return resultados_validacion
    
    def escribir_validacion_archivos(self, archivo_info):
        """Escribe la información de validación de archivos en el archivo de información"""
        try:
            resultados = self.validar_archivos_dependencias()
            
            with open(archivo_info, 'a', encoding='utf-8') as f:
                f.write(f"\nVALIDACIÓN DE ARCHIVOS DE DEPENDENCIAS:\n")
                f.write("="*60 + "\n")
                f.write(f"Fecha de validación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                for i, resultado in enumerate(resultados, 1):
                    f.write(f"{i}. {resultado['archivo']}\n")
                    f.write(f"   Ruta: {resultado['ruta_completa']}\n")
                    f.write(f"   Existe: {'SÍ' if resultado['existe'] else 'NO'}\n")
                    f.write(f"   Última modificación: {resultado['fecha_modificacion']}\n")
                    f.write(f"   Peso: {resultado['peso_mb']}\n")
                    f.write(f"   Días de antigüedad: {resultado['dias_diferencia']}\n")
                    f.write(f"   Estado: {resultado['estado']}\n")
                    f.write("-" * 40 + "\n")
                
                # Resumen general
                archivos_encontrados = sum(1 for r in resultados if r['existe'])
                archivos_recientes = sum(1 for r in resultados if r['existe'] and r['dias_diferencia'] != 'No disponible' and r['dias_diferencia'] <= 7)
                
                f.write(f"\nRESUMEN DE VALIDACIÓN:\n")
                f.write(f"Total de archivos verificados: {len(resultados)}\n")
                f.write(f"Archivos encontrados: {archivos_encontrados}\n")
                f.write(f"Archivos recientes (≤7 días): {archivos_recientes}\n")
                f.write(f"Estado general: {'ÓPTIMO' if archivos_encontrados == len(resultados) and archivos_recientes >= 2 else 'REVISAR ARCHIVOS ANTIGUOS' if archivos_encontrados == len(resultados) else 'ARCHIVOS FALTANTES'}\n")
                f.write("="*60 + "\n")
                
        except Exception as e:
            # Si hay error, al menos intentar escribir que hubo un problema
            try:
                with open(archivo_info, 'a', encoding='utf-8') as f:
                    f.write(f"\nERROR EN VALIDACIÓN DE ARCHIVOS: {str(e)}\n")
                    f.write("="*60 + "\n")
            except:
                pass
        
    def procesar_bom_completo(self):
        """Procesa el BOM completo key por key"""
        try:
            # Cargar datos
            if not self.bom_builder.cargar_datos_bom():
                raise Exception("No se pudieron cargar los datos")
            
            # Obtener todos los keys únicos
            keys_unicos = self.bom_builder.obtener_keys_unicos()
            
            if not keys_unicos:
                messagebox.showwarning("Advertencia", "No se encontraron keys para procesar")
                return
            
            # Crear ventana de progreso
            self.progress_window = ProgressWindow(self.parent, len(keys_unicos))
            
            # Configurar rutas de archivos (nombres fijos)
            script_dir = os.path.dirname(os.path.abspath(__file__))
            self.db_nueva_path = os.path.join(script_dir, "BOM_Procesado.db")
            self.log_file_path = os.path.join(script_dir, "BOM_Log.txt")
            self.info_file_path = os.path.join(script_dir, "BOM_Info.txt")
            
            # Verificar y eliminar base de datos existente
            if os.path.exists(self.db_nueva_path):
                try:
                    os.remove(self.db_nueva_path)
                    print(f"Base de datos existente eliminada: {self.db_nueva_path}")
                except Exception as e:
                    raise Exception(f"No se pudo eliminar la base de datos existente: {str(e)}")
            
            # Crear archivo de información con timestamp
            timestamp = datetime.now()
            with open(self.info_file_path, 'w', encoding='utf-8') as f:
                f.write(f"INFORMACIÓN DE CREACIÓN DE BASE DE DATOS BOM\n")
                f.write("="*50 + "\n")
                f.write(f"Fecha de creación: {timestamp.strftime('%Y-%m-%d')}\n")
                f.write(f"Hora de creación: {timestamp.strftime('%H:%M:%S')}\n")
                f.write(f"Timestamp completo: {timestamp}\n")
                f.write(f"Base de datos: {self.db_nueva_path}\n")
                f.write(f"Log de proceso: {self.log_file_path}\n")
                f.write("="*50 + "\n\n")
            
            # Escribir validación de archivos de dependencias
            self.escribir_validacion_archivos(self.info_file_path)
            
            # Inicializar log file
            with open(self.log_file_path, 'w', encoding='utf-8') as f:
                f.write(f"Log de Procesamiento BOM - {timestamp}\n")
                f.write("="*50 + "\n\n")
                f.write(f"Base de datos creada: {self.db_nueva_path}\n")
                f.write(f"Archivo de información: {self.info_file_path}\n\n")
            
            # Crear tabla en la nueva base de datos
            self.crear_tabla_nueva()
            
            # Procesar en un thread separado
            thread = threading.Thread(target=self.procesar_keys, args=(keys_unicos,))
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al iniciar el proceso:\n{str(e)}")
    
    def crear_tabla_nueva(self):
        """Crea la tabla en la nueva base de datos con todas las columnas"""
        conn = sqlite3.connect(self.db_nueva_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS bom_procesado (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                original_id INTEGER,
                Level_Number TEXT,
                Component TEXT,
                Description TEXT,
                CE TEXT,
                T TEXT,
                Sort TEXT,
                UM TEXT,
                MLI TEXT,
                Plan_Type TEXT,
                Unit_Qty REAL,
                Std_Cost REAL,
                Ext_Std REAL,
                Labor_IN REAL,
                Lab_Rem REAL,
                Mat_IN REAL,
                Mat_Rem REAL,
                This_lvl REAL,
                Lab_Hrs REAL,
                key TEXT,
                Orden_BOM_Original INTEGER,
                LT_in92 REAL,
                LT_PlanType REAL,
                LT_No_in_R4 REAL,
                LT_Final_parchado REAL,
                arbol TEXT,
                arbol_Plantype TEXT,
                arbol_LT TEXT,
                LT_sistema REAL,
                LT_Release REAL,
                fecha_procesamiento TIMESTAMP,
                key_procesado TEXT
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def procesar_keys(self, keys_unicos):
        """Procesa todos los keys uno por uno con lógica completa"""
        total_processed = 0
        total_errors = 0
        total_skipped = 0
        
        conn = sqlite3.connect(self.db_nueva_path)
        
        try:
            for i, key in enumerate(keys_unicos):
                # Verificar si está pausado
                while self.progress_window.paused and not self.progress_window.cancelled:
                    time.sleep(0.1)
                
                # Verificar si está cancelado
                if self.progress_window.cancelled:
                    self.progress_window.add_log("Proceso cancelado", "error")
                    break
                
                # Actualizar progreso
                self.progress_window.update_progress(i, key, total_processed, total_errors, total_skipped)
                
                retry = True
                while retry:
                    retry = False
                    try:
                        # Procesar el key
                        self.progress_window.add_log(f"Procesando key: {key}", "info")
                        
                        # Obtener datos del key
                        df_key = self.bom_builder.obtener_bom_por_key(key)
                        
                        if df_key is None or df_key.empty:
                            self.progress_window.add_log(f"Key {key} vacío o no encontrado", "warning")
                            total_skipped += 1
                            self.guardar_en_log(f"OMITIDO - Key {key}: Vacío o no encontrado")
                            break
                        
                        # Crear árbol con lógica completa
                        self.progress_window.add_log(f"Aplicando lógica completa a {key}...", "info")
                        df_procesado = self.bom_builder.crear_arbol_bom(df_key)
                        
                        # Añadir columnas adicionales
                        df_procesado['fecha_procesamiento'] = datetime.now()
                        df_procesado['key_procesado'] = key
                        
                        # Renombrar id a original_id para evitar conflictos
                        if 'id' in df_procesado.columns:
                            df_procesado = df_procesado.rename(columns={'id': 'original_id'})
                        
                        # Guardar en base de datos
                        df_procesado.to_sql('bom_procesado', conn, if_exists='append', index=False)
                        
                        total_processed += 1
                        
                        # Log detallado del procesamiento
                        registros_procesados = len(df_procesado)
                        registros_con_lt_release = len(df_procesado[df_procesado['LT_Release'] > 0])
                        registros_con_plantype = len(df_procesado[df_procesado['LT_PlanType'] > 0])
                        
                        self.progress_window.add_log(
                            f"Key {key} procesado: {registros_procesados} registros, "
                            f"{registros_con_lt_release} con LT_Release, "
                            f"{registros_con_plantype} con LT_PlanType", 
                            "success"
                        )
                        
                        self.guardar_en_log(
                            f"ÉXITO - Key {key}: {registros_procesados} registros procesados, "
                            f"{registros_con_lt_release} con LT_Release, "
                            f"{registros_con_plantype} con LT_PlanType"
                        )
                        
                    except Exception as e:
                        error_msg = str(e)
                        total_errors += 1
                        self.guardar_en_log(f"ERROR - Key {key}: {error_msg}\n{traceback.format_exc()}")
                        
                        if not self.progress_window.continue_on_error:
                            response = self.progress_window.show_error(error_msg, key)
                            
                            if response == "retry":
                                retry = True
                                self.progress_window.add_log(f"Reintentando key {key}", "warning")
                            elif response == "cancel":
                                self.progress_window.cancelled = True
                                break
                            # Si es "continue", simplemente continúa con el siguiente
                        else:
                            self.progress_window.add_log(f"Error en {key}: {error_msg} (continuando automáticamente)", "error")
            
            # Finalizar
            conn.commit()
            conn.close()
            
            self.progress_window.finish(total_processed, total_errors)
            
            # Guardar resumen en log
            self.guardar_resumen_final(total_processed, total_errors, total_skipped)
            
        except Exception as e:
            conn.close()
            self.progress_window.add_log(f"Error crítico: {str(e)}", "error")
            messagebox.showerror("Error Crítico", f"Error crítico en el proceso:\n{str(e)}")
    
    def guardar_en_log(self, mensaje):
        """Guarda un mensaje en el archivo de log"""
        try:
            with open(self.log_file_path, 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensaje}\n")
        except:
            pass
    
    def guardar_resumen_final(self, procesados, errores, omitidos):
        """Guarda el resumen final en el log y actualiza el archivo de información"""
        timestamp_final = datetime.now()
        
        # Resumen para el log
        resumen = f"""
{"="*50}
RESUMEN FINAL
{"="*50}
Fecha y hora de finalización: {timestamp_final}
Total de keys procesados: {procesados}
Total de errores: {errores}
Total de omitidos: {omitidos}
Base de datos generada: {self.db_nueva_path}

CARACTERÍSTICAS IMPLEMENTADAS:
- LT_Final_parchado: Combina LT_No_in_R4 y LT_in92
- Herencia LT_PlanType: Se propaga a niveles hijos
- LT_sistema: LT acumulado sin incluir nivel actual
- LT_Release: Con lógica MLI (L vs no-L) + 3 días preparación
- arbol: Jerarquía de componentes
- arbol_Plantype: Jerarquía de tipos de plan
- arbol_LT: LT recorrido acumulado por nivel
{"="*50}
"""
        self.guardar_en_log(resumen)
        
        # Actualizar archivo de información con resultados finales
        try:
            with open(self.info_file_path, 'a', encoding='utf-8') as f:
                f.write(f"\nRESULTADOS DEL PROCESAMIENTO:\n")
                f.write("="*50 + "\n")
                f.write(f"Fecha de finalización: {timestamp_final.strftime('%Y-%m-%d')}\n")
                f.write(f"Hora de finalización: {timestamp_final.strftime('%H:%M:%S')}\n")
                f.write(f"Total de keys procesados: {procesados}\n")
                f.write(f"Total de errores: {errores}\n")
                f.write(f"Total de omitidos: {omitidos}\n")
                f.write(f"Estado: {'COMPLETADO CON ERRORES' if errores > 0 else 'COMPLETADO EXITOSAMENTE'}\n")
                f.write("="*50 + "\n")
        except:
            pass

class BOMApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador BOM con Lógica Completa")
        self.root.geometry("800x600")
        
        # Centrar la ventana
        self.centrar_ventana()
        
        self.bom_processor = BOMProcessor(root)
        
        self.crear_interfaz()
        
    def centrar_ventana(self):
        """Centra la ventana en la pantalla"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def crear_interfaz(self):
        """Crea la interfaz gráfica simplificada"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar el grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="Procesador BOM con Lógica Completa", 
                          font=('Arial', 16, 'bold'))
        titulo.grid(row=0, column=0, pady=(0, 20))
        
        # Frame para información principal
        info_frame = ttk.LabelFrame(main_frame, text="Procesamiento:", padding="15")
        info_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10)
        info_frame.columnconfigure(0, weight=1)
        
        # Descripción principal
        descripcion = """Este procesador aplicará lógica completa a todo el BOM:

• Procesa KEY por KEY con toda la lógica implementada
• Crea nueva base de datos SQLite con resultados
• Permite pausar/reanudar el proceso
• Genera log detallado de errores
• Manejo inteligente de errores con opciones de recuperación"""
        
        ttk.Label(info_frame, text=descripcion, justify=tk.LEFT).grid(row=0, column=0, sticky=tk.W)
        
        # Frame para características implementadas
        features_frame = ttk.LabelFrame(main_frame, text="Lógica Completa Implementada:", padding="15")
        features_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10)
        
        features_text = """✓ LT_Final_parchado: LT_No_in_R4 si ≠ 0, sino LT_in92
✓ Herencia LT_PlanType: Se propaga automáticamente a niveles hijos
✓ LT_sistema: Suma LT_Final_parchado de padres (SIN nivel actual)
✓ LT_Release: Si LT_PlanType > 0 → LT_PlanType + 3
             Si MLI = 'L' → LT_recorrido + 3 (material comprado)
             Si MLI ≠ 'L' → LT_recorrido + LT_nivel + 3 (material fabricado)
✓ arbol_LT: Muestra LT recorrido acumulado por nivel jerárquico
✓ arbol: Jerarquía de componentes
✓ arbol_Plantype: Jerarquía de tipos de plan"""
        
        ttk.Label(features_frame, text=features_text, justify=tk.LEFT, foreground="darkgreen").grid(row=0, column=0)
        
        # Frame para botones
        botones_frame = ttk.Frame(main_frame)
        botones_frame.grid(row=3, column=0, pady=30)
        
        # Botón Procesar BOM Completo
        self.btn_procesar = ttk.Button(botones_frame, text="Procesar BOM Completo", 
                                      command=self.procesar_bom,
                                      width=25)
        self.btn_procesar.grid(row=0, column=0, padx=5)
        
        # Botón Salir
        btn_salir = ttk.Button(botones_frame, text="Salir", 
                              command=self.root.quit,
                              width=15)
        btn_salir.grid(row=0, column=1, padx=5)
        
        # Label de estado
        self.lbl_estado = ttk.Label(main_frame, text="Listo para procesar BOM con lógica completa", 
                                   foreground="green", font=('Arial', 10, 'bold'))
        self.lbl_estado.grid(row=4, column=0, pady=10)
    
    def procesar_bom(self):
        """Procesa el BOM con lógica completa"""
        self.bom_processor.procesar_bom_completo()

def main():
    """Función principal para ejecutar la aplicación"""
    root = tk.Tk()
    app = BOMApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()