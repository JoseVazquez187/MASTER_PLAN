import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import pandas as pd
from datetime import datetime
import os

class ExpediteTableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Expedite Table Manager")
        self.root.geometry("900x700")
        
        # Variables de configuración
        self.path_conn = "BOM_Procesado.db"  # Base de datos existente
        self.folder_db = ""  # Se establecerá cuando selecciones el archivo
        
        # Ruta del archivo de días festivos
        self.holidays_path = r"C:\Users\J.Vazquez\Desktop\PARCHE EXPEDITE\Holidays.xlsx"
        
        self.setup_ui()
        
    def setup_ui(self):
        # Notebook para las pestañas
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Pestaña 1: Expedite Table Manager
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text="Expedite Manager")
        self.setup_expedite_tab(tab1)
        
        # Pestaña 2: Cruce de Tablas
        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text="Cruce Expedite vs BOM")
        self.setup_cruce_tab(tab2)
        
    def setup_expedite_tab(self, parent):
        # Frame principal
        main_frame = ttk.Frame(parent, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Expedite Table Manager", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Botón para seleccionar archivo CSV
        self.select_file_btn = ttk.Button(main_frame, text="Seleccionar Archivo Expedite.csv", 
                                         command=self.select_file)
        self.select_file_btn.grid(row=1, column=0, columnspan=3, pady=(0, 10))
        
        # Label para mostrar archivo seleccionado
        self.file_label = ttk.Label(main_frame, text="No se ha seleccionado archivo", 
                                   foreground="gray")
        self.file_label.grid(row=2, column=0, columnspan=3, pady=(0, 20))
        
        # Botón para crear tabla expedite
        self.create_btn = ttk.Button(main_frame, text="Crear Tabla Expedite", 
                                    command=self.create_expedite_table, state="disabled")
        self.create_btn.grid(row=3, column=0, columnspan=3, pady=(0, 20))
        
        # Label de estado (equivalente a label_23)
        self.label_23 = ttk.Label(main_frame, text="Procesando...", foreground="orange")
        self.label_23.grid(row=4, column=0, columnspan=3)
        self.label_23.grid_remove()  # Ocultar inicialmente
        
        # Label de fecha de actualización (equivalente a label_2)
        self.label_2 = ttk.Label(main_frame, text="", foreground="green")
        self.label_2.grid(row=5, column=0, columnspan=3, pady=(10, 0))
        self.label_2.grid_remove()  # Ocultar inicialmente
        
        # Frame para mostrar información de la tabla
        info_frame = ttk.LabelFrame(main_frame, text="Información de la Tabla", padding="10")
        info_frame.grid(row=6, column=0, columnspan=3, pady=(20, 0), sticky=(tk.W, tk.E))
        
        # Text widget para mostrar información
        self.info_text = tk.Text(info_frame, height=10, width=70, state="disabled")
        scrollbar = ttk.Scrollbar(info_frame, orient="vertical", command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        
        self.info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar el grid para que se expanda
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        info_frame.columnconfigure(0, weight=1)
        
    def setup_cruce_tab(self, parent):
        # Frame principal para el cruce
        main_frame = ttk.Frame(parent, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Cruce Expedite vs BOM Procesado", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Frame de información
        info_frame = ttk.LabelFrame(main_frame, text="Información del Proceso", padding="10")
        info_frame.grid(row=1, column=0, columnspan=3, pady=(0, 20), sticky=(tk.W, tk.E))
        
        info_text = tk.Text(info_frame, height=8, width=80, state="disabled")
        info_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        info_content = """PROCESO DE CRUCE DE TABLAS CON VALIDACIONES:

1. Validaciones implementadas:
   ✅ Estandarización de llaves (UPPER y TRIM)
   ✅ Eliminación de duplicados por LT_Release mayor
   ✅ Verificación de integridad (todas las filas de expedite procesadas)
   ✅ Detección de registros duplicados

2. Lógica de cruce por prioridad:
   - BOM_PROCESADO: UPPER(TRIM(key + Component + Sort))
   - EXPEDITE: UPPER(TRIM(DemandSource + ItemNo + Sort))
   - Si no hay match pero DemandSource = ItemNo → usar LT_Release = 3

3. Resultado: Tabla EXPEDITE_PARCHADO con métricas completas
   - Backschedule con BOM (LT_Release del BOM)
   - Backschedule regla especial (LT_Release = 3)
   - Sin backschedule (fecha original estandarizada)"""
        
        info_text.config(state="normal")
        info_text.insert(1.0, info_content)
        info_text.config(state="disabled")
        
        # Botón para ejecutar cruce
        self.cruce_btn = ttk.Button(main_frame, text="Ejecutar Cruce y Backschedule", 
                                   command=self.ejecutar_cruce, style="Accent.TButton")
        self.cruce_btn.grid(row=2, column=0, columnspan=3, pady=(0, 10))
        
        # Barra de progreso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                          maximum=100, length=300)
        self.progress_bar.grid(row=3, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))
        
        # Labels de estado para el cruce
        self.cruce_status_label = ttk.Label(main_frame, text="", foreground="orange")
        self.cruce_status_label.grid(row=4, column=0, columnspan=3)
        
        self.cruce_result_label = ttk.Label(main_frame, text="", foreground="green")
        self.cruce_result_label.grid(row=5, column=0, columnspan=3, pady=(10, 0))
        
        # Frame para mostrar resultados del cruce
        result_frame = ttk.LabelFrame(main_frame, text="Resultados del Cruce", padding="10")
        result_frame.grid(row=6, column=0, columnspan=3, pady=(20, 0), sticky=(tk.W, tk.E))
        
        self.cruce_info_text = tk.Text(result_frame, height=12, width=80, state="disabled")
        scrollbar2 = ttk.Scrollbar(result_frame, orient="vertical", command=self.cruce_info_text.yview)
        self.cruce_info_text.configure(yscrollcommand=scrollbar2.set)
        
        self.cruce_info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar2.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar el grid
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)
        
    def select_file(self):
        """Seleccionar archivo CSV"""
        file_path = filedialog.askopenfilename(
            title="Seleccionar Expedite.csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.folder_db = os.path.dirname(file_path) + "/"
            self.csv_filename = os.path.basename(file_path)
            self.file_label.config(text=f"Archivo: {self.csv_filename}", foreground="blue")
            self.create_btn.config(state="normal")
        
    def create_expedite_table(self):
        """Crear tabla expedite - equivalente al método original"""
        self.label_23.grid()  # Mostrar label de procesando
        
        try:
            self.label_2.grid()  # Mostrar label de fecha
            
            # Verificar que la base de datos existe
            if not os.path.exists(self.path_conn):
                messagebox.showerror("Error", f"No se encontró la base de datos: {self.path_conn}")
                self.label_2.grid_remove()
                return
            
            # Conectar a la base de datos existente
            conn = sqlite3.connect(self.path_conn)
            cursor = conn.cursor()
            
            # Verificar si la tabla existe y eliminar contenido
            cursor.execute("""SELECT name FROM sqlite_master WHERE type='table' AND name='expedite'""")
            table_exists = cursor.fetchone()
            
            if table_exists:
                cursor.execute("""DELETE FROM expedite""")
                conn.commit()
                print("Contenido de tabla expedite eliminado")
            
            # Crear tabla si no existe
            cursor.execute("""CREATE TABLE IF NOT EXISTS expedite(id INTEGER PRIMARY KEY
                        , EntityGroup TEXT
                        , Project TEXT
                        , AC TEXT
                        , ItemNo TEXT
                        , Description TEXT
                        , PlanTp TEXT
                        , Ref TEXT
                        , Sub TEXT
                        , FillDoc TEXT
                        , DemandType TEXT
                        , Sort TEXT, ReqQty TEXT, DemandSource TEXT, Unit TEXT
                        , Vendor TEXT, ReqDate TEXT, ShipDate TEXT, OH TEXT, MLIKCode TEXT, LT TEXT
                        , STDCost TEXT
                        , LotSize TEXT
                        , UOM TEXT)""")
            conn.commit()
            print("Tabla expedite creada/verificada")
            
            try:
                # Leer archivo CSV (saltando la primera fila ya que los headers están en row 2)
                csv_path = os.path.join(self.folder_db, self.csv_filename)
                import_file = pd.read_csv(csv_path, skiprows=1)  # Saltar primera fila, headers en row 2
                
                # Seleccionar columnas específicas
                try:
                    import_file = import_file[['Entity Group','Project','A/C','Item-No', 'Description'
                                            , 'PlanTp','Ref','Sub','Sort','Fill-Doc','Demand - Type','Req-Qty',
                                            'Demand-Source','Unit','Vendor','Req-Date','Ship-Date','OH','MLIK Code','LT','Std-Cost','Lot-Size','UOM']]
                except KeyError as e:
                    messagebox.showerror("Error de Columnas", 
                                       f"No se encontraron las columnas esperadas: {e}\n\n"
                                       f"Columnas disponibles: {list(import_file.columns)}")
                    return
                
                # Renombrar columnas
                import_file = import_file.rename(columns = {
                    'Entity Group':'EntityGroup',
                    'A/C':'AC',
                    'Item-No':'ItemNo',
                    'Req-Qty':'ReqQty',
                    'Demand-Source':'DemandSource',
                    'Req-Date':'ReqDate',
                    'Ship-Date':'ShipDate',
                    'MLIK Code':'MLIKCode',
                    'Std-Cost':'STDCost',
                    'Lot-Size':'LotSize',
                    'Fill-Doc':'FillDoc',
                    'Demand - Type':'DemandType'
                })
                
                # Insertar datos en la base de datos
                import_file.to_sql('expedite', conn, if_exists='append', index=False)
                conn.commit()
                print(f"Datos insertados: {len(import_file)} registros")
                
                # Actualizar fecha
                now = datetime.now()
                today = now.strftime("%m/%d/%Y %I:%M:%S")
                self.label_2.config(text=f'Última actualización: {today}')
                
                # Mostrar información de la tabla
                self.show_table_info(import_file)
                
                messagebox.showinfo("Éxito", f"Tabla expedite actualizada exitosamente!\n"
                                           f"Base de datos: {self.path_conn}\n"
                                           f"Registros procesados: {len(import_file)}\n"
                                           f"Fecha: {today}")
                
            except sqlite3.Error as error:
                messagebox.showerror("Error en base Expedite", f'No se pudo establecer conexión: {error}')
                self.label_2.grid_remove()
            except Exception as error:
                messagebox.showerror("Error", f'Error al procesar el archivo: {error}')
                self.label_2.grid_remove()
            finally:
                conn.close()
                
        except sqlite3.Error as error:
            messagebox.showerror("Error en base de datos", f'No se pudo establecer conexión: {error}')
            self.label_2.grid_remove()
        finally:
            self.label_23.grid_remove()  # Ocultar label de procesando
    
    def show_table_info(self, dataframe):
        """Mostrar información de la tabla procesada"""
        self.info_text.config(state="normal")
        self.info_text.delete(1.0, tk.END)
        
        info = f"=== INFORMACIÓN DE LA TABLA EXPEDITE ===\n\n"
        info += f"Registros totales: {len(dataframe)}\n"
        info += f"Columnas: {len(dataframe.columns)}\n\n"
        
        info += "Columnas procesadas:\n"
        for i, col in enumerate(dataframe.columns, 1):
            info += f"{i:2d}. {col}\n"
        
        info += f"\nPrimeros 5 registros:\n"
        info += "=" * 50 + "\n"
        info += dataframe.head().to_string()
        
        info += f"\n\n=== ESTADÍSTICAS ===\n"
        info += f"Valores únicos por columna:\n"
        for col in dataframe.columns:
            unique_count = dataframe[col].nunique()
            info += f"  {col}: {unique_count} valores únicos\n"
        
        self.info_text.insert(1.0, info)
        self.info_text.config(state="disabled")
        
    def load_holidays(self):
        """Cargar días festivos desde el archivo Excel"""
        try:
            holidays_df = pd.read_excel(self.holidays_path)
            # Convertir a datetime y extraer solo las fechas
            holidays_dates = pd.to_datetime(holidays_df.iloc[:, 0]).dt.date.tolist()
            return holidays_dates
        except Exception as e:
            print(f"Error cargando días festivos: {e}")
            return []
    
    def standardize_key(self, key_value):
        """Estandarizar llave: UPPER y TRIM"""
        if pd.isna(key_value):
            return ""
        return str(key_value).strip().upper()
    
    def calculate_business_days_back(self, end_date, business_days, holidays):
        """Calcular fecha hacia atrás considerando solo días laborales y festivos"""
        if pd.isna(end_date) or business_days == 0:
            return end_date
            
        try:
            # Convertir a datetime si es string
            if isinstance(end_date, str):
                end_date = pd.to_datetime(end_date)
            elif not isinstance(end_date, pd.Timestamp):
                end_date = pd.to_datetime(end_date)
            
            current_date = end_date
            days_counted = 0
            
            while days_counted < business_days:
                current_date = current_date - pd.Timedelta(days=1)
                
                # Verificar si es día laboral (lunes=0 a viernes=4)
                if current_date.weekday() < 5 and current_date.date() not in holidays:
                    days_counted += 1
            
            # Retornar en formato estándar YYYY-MM-DD
            return current_date.strftime('%Y-%m-%d')
            
        except Exception as e:
            print(f"Error calculando backschedule: {e}")
            # Si hay error, intentar convertir la fecha original al formato estándar
            try:
                if pd.notna(end_date):
                    return pd.to_datetime(end_date).strftime('%Y-%m-%d')
                else:
                    return None
            except:
                return end_date
    
    def create_bom_index_with_deduplication(self, conn):
        """Crear índice de BOM eliminando duplicados por LT_Release mayor"""
        self.cruce_status_label.config(text="Creando índice de BOM con deduplicación...")
        self.root.update()
        
        # Leer BOM con las columnas necesarias
        bom_query = """
        SELECT key, Component, Sort, LT_Release, id
        FROM bom_procesado
        WHERE LT_Release IS NOT NULL 
        AND key IS NOT NULL 
        AND Component IS NOT NULL 
        AND Sort IS NOT NULL
        """
        bom_df = pd.read_sql_query(bom_query, conn)
        
        # Estandarizar y crear llave de cruce
        bom_df['cruce_key'] = bom_df.apply(
            lambda row: self.standardize_key(row['key']) + 
                       self.standardize_key(row['Component']) + 
                       self.standardize_key(row['Sort']), axis=1
        )
        
        # Convertir LT_Release a numérico
        bom_df['LT_Release'] = pd.to_numeric(bom_df['LT_Release'], errors='coerce')
        
        # Eliminar duplicados manteniendo el LT_Release mayor
        bom_dedup = bom_df.sort_values('LT_Release', ascending=False).drop_duplicates(subset=['cruce_key'], keep='first')
        
        # Crear diccionario para búsqueda rápida
        bom_dict = dict(zip(bom_dedup['cruce_key'], bom_dedup['LT_Release']))
        
        # Métricas de deduplicación
        original_count = len(bom_df)
        deduplicated_count = len(bom_dedup)
        duplicates_removed = original_count - deduplicated_count
        
        print(f"BOM original: {original_count} registros")
        print(f"BOM deduplicado: {deduplicated_count} registros")
        print(f"Duplicados eliminados: {duplicates_removed}")
        
        return bom_dict, {
            'original_count': original_count,
            'deduplicated_count': deduplicated_count,
            'duplicates_removed': duplicates_removed
        }
    
    def validate_integrity(self, conn):
        """Validar integridad: todas las filas de expedite deben estar en expedite_parchado"""
        self.cruce_status_label.config(text="Validando integridad de datos...")
        self.root.update()
        
        # Contar registros en expedite
        expedite_count = pd.read_sql_query("SELECT COUNT(*) as count FROM expedite", conn).iloc[0]['count']
        
        # Contar registros en expedite_parchado
        parchado_count = pd.read_sql_query("SELECT COUNT(*) as count FROM expedite_parchado", conn).iloc[0]['count']
        
        # Verificar duplicados en expedite_parchado
        duplicates_query = """
        SELECT cruce_key, COUNT(*) as count
        FROM (
            SELECT 
                UPPER(TRIM(DemandSource || ItemNo || Sort)) as cruce_key
            FROM expedite_parchado
        )
        GROUP BY cruce_key
        HAVING COUNT(*) > 1
        """
        duplicates_df = pd.read_sql_query(duplicates_query, conn)
        duplicates_count = len(duplicates_df)
        total_duplicate_rows = duplicates_df['count'].sum() - duplicates_count if duplicates_count > 0 else 0
        
        return {
            'expedite_count': expedite_count,
            'parchado_count': parchado_count,
            'missing_records': max(0, expedite_count - parchado_count),
            'duplicates_keys': duplicates_count,
            'total_duplicate_rows': total_duplicate_rows,
            'integrity_ok': expedite_count == parchado_count and duplicates_count == 0
        }
    
    def ejecutar_cruce(self):
        """Ejecutar el cruce entre expedite y bom_procesado por lotes con validaciones completas"""
        self.cruce_btn.config(state="disabled")
        self.progress_var.set(0)
        self.cruce_status_label.config(text="Iniciando proceso de cruce...")
        self.root.update()
        
        metrics = {
            'start_time': datetime.now(),
            'bom_metrics': {},
            'processing_metrics': {},
            'validation_metrics': {},
            'end_time': None
        }
        
        try:
            # Verificar que la base de datos existe
            if not os.path.exists(self.path_conn):
                messagebox.showerror("Error", f"No se encontró la base de datos: {self.path_conn}")
                return
            
            conn = sqlite3.connect(self.path_conn)
            
            # 1. Cargar días festivos
            self.cruce_status_label.config(text="Cargando días festivos...")
            self.progress_var.set(5)
            self.root.update()
            holidays = self.load_holidays()
            print(f"Días festivos cargados: {len(holidays)}")
            
            # 2. Obtener conteo total de registros de expedite
            total_count_query = "SELECT COUNT(*) FROM expedite"
            total_records = pd.read_sql_query(total_count_query, conn).iloc[0, 0]
            print(f"Total registros expedite: {total_records}")
            
            # 3. Crear índice de BOM con deduplicación
            self.progress_var.set(10)
            bom_dict, bom_metrics = self.create_bom_index_with_deduplication(conn)
            metrics['bom_metrics'] = bom_metrics
            
            # 4. Crear tabla expedite_parchado
            self.cruce_status_label.config(text="Creando tabla expedite_parchado...")
            self.progress_var.set(15)
            self.root.update()
            
            cursor = conn.cursor()
            cursor.execute("DROP TABLE IF EXISTS expedite_parchado")
            
            # Crear estructura de la tabla
            create_table_query = """
            CREATE TABLE expedite_parchado (
                id INTEGER,
                EntityGroup TEXT,
                Project TEXT,
                AC TEXT,
                ItemNo TEXT,
                Description TEXT,
                PlanTp TEXT,
                Ref TEXT,
                Sub TEXT,
                FillDoc TEXT,
                DemandType TEXT,
                Sort TEXT,
                ReqQty TEXT,
                DemandSource TEXT,
                Unit TEXT,
                Vendor TEXT,
                ReqDate_Original TEXT,
                ShipDate TEXT,
                OH TEXT,
                MLIKCode TEXT,
                LT TEXT,
                STDCost TEXT,
                LotSize TEXT,
                UOM TEXT,
                LT_Release TEXT,
                ReqDate TEXT,
                cruce_key TEXT
            )
            """
            cursor.execute(create_table_query)
            conn.commit()
            
            # 5. Procesar por lotes
            chunk_size = 1000
            offset = 0
            processed_records = 0
            matches = 0
            no_matches = 0
            bom_matches = 0
            demandsource_itemno_matches = 0
            key_standardization_issues = 0
            
            while offset < total_records:
                # Leer chunk de expedite
                expedite_query = f"""
                SELECT id, EntityGroup, Project, AC, ItemNo, Description, PlanTp, Ref, Sub, 
                       FillDoc, DemandType, Sort, ReqQty, DemandSource, Unit, Vendor, 
                       ReqDate, ShipDate, OH, MLIKCode, LT, STDCost, LotSize, UOM
                FROM expedite
                LIMIT {chunk_size} OFFSET {offset}
                """
                
                chunk_df = pd.read_sql_query(expedite_query, conn)
                
                if len(chunk_df) == 0:
                    break
                
                # Actualizar progreso
                progress = 20 + (offset / total_records) * 65  # 20% a 85%
                self.progress_var.set(progress)
                self.cruce_status_label.config(text=f"Procesando registros {offset + 1} a {offset + len(chunk_df)} de {total_records}")
                self.root.update()
                
                # Procesar chunk
                processed_chunk = []
                
                for _, row in chunk_df.iterrows():
                    # Crear llave de cruce estandarizada para expedite
                    try:
                        cruce_key = (self.standardize_key(row['DemandSource']) + 
                                   self.standardize_key(row['ItemNo']) + 
                                   self.standardize_key(row['Sort']))
                    except Exception as e:
                        key_standardization_issues += 1
                        cruce_key = ""
                        print(f"Error estandarizando llave: {e}")
                    
                    # Buscar LT_Release en el diccionario
                    lt_release = bom_dict.get(cruce_key, None)
                    
                    # Calcular nueva ReqDate
                    new_req_date = None
                    lt_used = None
                    match_type = None
                    
                    if lt_release is not None:
                        # Caso 1: Match encontrado en BOM
                        matches += 1
                        bom_matches += 1
                        lt_used = lt_release
                        match_type = "BOM_MATCH"
                        new_req_date = self.calculate_business_days_back(
                            row['ShipDate'], 
                            int(lt_release),
                            holidays
                        )
                    elif str(row['DemandSource']).strip().upper() == str(row['ItemNo']).strip().upper():
                        # Caso 2: No hay match en BOM pero DemandSource = ItemNo, usar LT=3
                        matches += 1  # Contar como match porque aplicamos regla
                        demandsource_itemno_matches += 1
                        lt_used = 3
                        match_type = "DEMANDSOURCE_ITEMNO_MATCH"
                        new_req_date = self.calculate_business_days_back(
                            row['ShipDate'], 
                            3,
                            holidays
                        )
                    else:
                        # Caso 3: Sin match, mantener fecha original
                        no_matches += 1
                        lt_used = None
                        match_type = "NO_MATCH"
                        # Convertir fecha original al formato estándar YYYY-MM-DD
                        original_date = row['ReqDate']
                        if pd.notna(original_date):
                            try:
                                # Convertir a datetime y luego a formato estándar
                                if isinstance(original_date, str):
                                    parsed_date = pd.to_datetime(original_date)
                                else:
                                    parsed_date = pd.to_datetime(original_date)
                                new_req_date = parsed_date.strftime('%Y-%m-%d')
                            except:
                                new_req_date = original_date  # Si falla, mantener original
                        else:
                            new_req_date = None
                    
                    # Preparar registro para insertar
                    processed_row = (
                        row['id'], row['EntityGroup'], row['Project'], row['AC'], row['ItemNo'],
                        row['Description'], row['PlanTp'], row['Ref'], row['Sub'], row['FillDoc'],
                        row['DemandType'], row['Sort'], row['ReqQty'], row['DemandSource'], row['Unit'],
                        row['Vendor'], row['ReqDate'], row['ShipDate'], row['OH'], row['MLIKCode'],
                        row['LT'], row['STDCost'], row['LotSize'], row['UOM'], lt_used, new_req_date, cruce_key
                    )
                    processed_chunk.append(processed_row)
                
                # Insertar chunk en la base de datos
                insert_query = """
                INSERT INTO expedite_parchado (
                    id, EntityGroup, Project, AC, ItemNo, Description, PlanTp, Ref, Sub, FillDoc,
                    DemandType, Sort, ReqQty, DemandSource, Unit, Vendor, ReqDate_Original, ShipDate,
                    OH, MLIKCode, LT, STDCost, LotSize, UOM, LT_Release, ReqDate, cruce_key
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
                
                cursor.executemany(insert_query, processed_chunk)
                conn.commit()
                
                processed_records += len(chunk_df)
                offset += chunk_size
                
                # Liberar memoria
                del chunk_df, processed_chunk
            
            # 6. Guardar métricas de procesamiento
            metrics['processing_metrics'] = {
                'total_processed': processed_records,
                'matches': matches,
                'bom_matches': bom_matches,
                'demandsource_itemno_matches': demandsource_itemno_matches,
                'no_matches': no_matches,
                'key_standardization_issues': key_standardization_issues,
                'match_rate': (matches / processed_records * 100) if processed_records > 0 else 0,
                'bom_match_rate': (bom_matches / processed_records * 100) if processed_records > 0 else 0,
                'demandsource_match_rate': (demandsource_itemno_matches / processed_records * 100) if processed_records > 0 else 0
            }
            
            # 7. Validar integridad
            self.progress_var.set(90)
            validation_metrics = self.validate_integrity(conn)
            metrics['validation_metrics'] = validation_metrics
            
            # 8. Finalizar
            metrics['end_time'] = datetime.now()
            processing_time = (metrics['end_time'] - metrics['start_time']).total_seconds()
            
            self.progress_var.set(100)
            self.cruce_status_label.config(text="Proceso completado!")
            
            now = datetime.now()
            today = now.strftime("%m/%d/%Y %I:%M:%S")
            self.cruce_result_label.config(text=f'Cruce completado: {today}')
            
            # Mostrar resultados finales con métricas completas
            self.show_complete_metrics(metrics, len(holidays), processing_time)
            
            # Determinar el estado del proceso
            status = "✅ EXITOSO" if validation_metrics['integrity_ok'] else "⚠️ CON ADVERTENCIAS"
            
            messagebox.showinfo("Cruce Completado", 
                              f"Proceso {status}\n\n"
                              f"📊 RESUMEN:\n"
                              f"Total procesados: {processed_records:,}\n"
                              f"BOM matches: {bom_matches:,} ({metrics['processing_metrics']['bom_match_rate']:.1f}%)\n"
                              f"DemandSource=ItemNo: {demandsource_itemno_matches:,} ({metrics['processing_metrics']['demandsource_match_rate']:.1f}%)\n"
                              f"Sin match: {no_matches:,}\n"
                              f"Total con backschedule: {matches:,} ({metrics['processing_metrics']['match_rate']:.1f}%)\n\n"
                              f"🔍 VALIDACIONES:\n"
                              f"Integridad: {'✅ OK' if validation_metrics['integrity_ok'] else '⚠️ Revisar'}\n"
                              f"Duplicados BOM eliminados: {bom_metrics['duplicates_removed']:,}\n"
                              f"Tiempo: {processing_time:.1f}s")
            
        except Exception as error:
            messagebox.showerror("Error en Cruce", f'Error durante el cruce: {error}')
            print(f"Error detallado: {error}")
        finally:
            if 'conn' in locals():
                conn.close()
            self.cruce_btn.config(state="normal")
            self.progress_var.set(0)
    
    def show_complete_metrics(self, metrics, holidays_count, processing_time):
        """Mostrar métricas completas del cruce realizado"""
        self.cruce_info_text.config(state="normal")
        self.cruce_info_text.delete(1.0, tk.END)
        
        bom = metrics['bom_metrics']
        proc = metrics['processing_metrics']
        val = metrics['validation_metrics']
        
        info = f"=== 📊 MÉTRICAS COMPLETAS DEL PROCESO ===\n\n"
        
        # Información general
        info += f"🕒 TIEMPO DE PROCESAMIENTO\n"
        info += f"Inicio: {metrics['start_time'].strftime('%H:%M:%S')}\n"
        info += f"Fin: {metrics['end_time'].strftime('%H:%M:%S')}\n"
        info += f"Duración: {processing_time:.1f} segundos\n"
        info += f"Velocidad: {proc['total_processed']/processing_time:.0f} registros/segundo\n\n"
        
        # Métricas de BOM
        info += f"🗃️ DEDUPLICACIÓN BOM_PROCESADO\n"
        info += f"Registros originales: {bom['original_count']:,}\n"
        info += f"Registros únicos: {bom['deduplicated_count']:,}\n"
        info += f"Duplicados eliminados: {bom['duplicates_removed']:,}\n"
        info += f"Tasa de duplicación: {(bom['duplicates_removed']/bom['original_count']*100):.1f}%\n\n"
        
        # Métricas de procesamiento
        info += f"⚙️ PROCESAMIENTO EXPEDITE\n"
        info += f"Total registros: {proc['total_processed']:,}\n"
        info += f"Matches BOM: {proc['bom_matches']:,} ({proc['bom_match_rate']:.1f}%)\n"
        info += f"Matches DemandSource=ItemNo (LT=3): {proc['demandsource_itemno_matches']:,} ({proc['demandsource_match_rate']:.1f}%)\n"
        info += f"Total con backschedule: {proc['matches']:,} ({proc['match_rate']:.1f}%)\n"
        info += f"Sin backschedule: {proc['no_matches']:,}\n"
        info += f"Problemas estandarización: {proc['key_standardization_issues']}\n\n"
        
        # Métricas de validación
        info += f"✅ VALIDACIÓN DE INTEGRIDAD\n"
        info += f"Registros expedite: {val['expedite_count']:,}\n"
        info += f"Registros expedite_parchado: {val['parchado_count']:,}\n"
        info += f"Registros faltantes: {val['missing_records']}\n"
        info += f"Llaves duplicadas: {val['duplicates_keys']}\n"
        info += f"Filas duplicadas totales: {val['total_duplicate_rows']}\n"
        info += f"Estado integridad: {'✅ CORRECTO' if val['integrity_ok'] else '⚠️ REVISAR'}\n\n"
        
        # Configuración del proceso
        info += f"🔧 CONFIGURACIÓN APLICADA\n"
        info += f"Días festivos considerados: {holidays_count}\n"
        info += f"Tamaño de lotes: 1,000 registros\n"
        info += f"Estandarización de llaves: UPPER + TRIM\n"
        info += f"Criterio duplicados BOM: LT_Release mayor\n"
        info += f"Formato fechas: YYYY-MM-DD\n\n"
        
        # Estructura de la tabla resultante
        info += f"📋 TABLA EXPEDITE_PARCHADO\n"
        info += f"Columnas totales: 27\n"
        info += f"- Todas las columnas de expedite\n"
        info += f"- LT_Release (del BOM)\n"
        info += f"- ReqDate_Original (fecha original)\n"
        info += f"- ReqDate (fecha calculada con backschedule)\n"
        info += f"- cruce_key (llave de cruce estandarizada)\n\n"
        
        # Recomendaciones
        if not val['integrity_ok']:
            info += f"⚠️ RECOMENDACIONES\n"
            if val['missing_records'] > 0:
                info += f"• Revisar {val['missing_records']} registros faltantes\n"
            if val['duplicates_keys'] > 0:
                info += f"• Investigar {val['duplicates_keys']} llaves duplicadas\n"
            if proc['key_standardization_issues'] > 0:
                info += f"• Verificar {proc['key_standardization_issues']} problemas de llaves\n"
        else:
            info += f"🎉 PROCESO COMPLETADO EXITOSAMENTE\n"
            info += f"Todos los controles de calidad pasaron correctamente.\n"
        
        info += f"\nFecha reporte: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        self.cruce_info_text.insert(1.0, info)
        self.cruce_info_text.config(state="disabled")

def main():
    root = tk.Tk()
    app = ExpediteTableApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()