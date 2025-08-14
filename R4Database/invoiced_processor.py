#!/usr/bin/env python3
"""
Procesador Especializado para tabla INVOICED - VERSIÓN CORREGIDA
Soporta modo INCREMENTAL (diario) y modo COMPLETO (histórico + actual)
"""

import os
import pandas as pd
import sqlite3
from datetime import datetime
from typing import List, Dict, Tuple, Any, Optional
import re

class InvoicedProcessor:
    """Procesador especializado para tabla invoiced con modo dual"""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.current_file = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\Invoiced\invExp.xlsx"
        self.historical_dir = r"J:\Departments\Operations\Shared\IT Administration\Python\IRPT\WHS PLAN\FILES\Invoiced\HISTORICO_YEAR"
        
        # ✅ FORZAR CREACIÓN DE TABLA AL INICIALIZAR
        self.ensure_table_exists()
    
    def ensure_table_exists(self):
        """Asegura que la tabla invoiced exista en la base de datos"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Crear la tabla invoiced - ESTRUCTURA SIMPLIFICADA
            cursor.execute("""CREATE TABLE IF NOT EXISTS invoiced(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Entity_Code TEXT, Invoice_No TEXT, Line TEXT, SO_No TEXT,
                Customer TEXT, Type TEXT, Cust_Name TEXT, PO_No TEXT, AC TEXT,
                Spr_C TEXT, O_Cd TEXT, Proj TEXT, TC TEXT, Item_Number TEXT,
                Description TEXT, UM TEXT, S_Qty REAL, Price REAL, Total REAL,
                Curr TEXT, CustReqDt TEXT, Req_Date TEXT, Pln_Ship TEXT,
                Inv_Dt TEXT, Inv_Time TEXT, Ord_Dt TEXT, SODueDt TEXT,
                Due_Dt TEXT, TimeToConfirm_hr REAL, Plan_FillRate REAL,
                TAT_to_fill_an_order REAL, Cust_FillRate REAL, SH TEXT,
                DG TEXT, Std_Cost REAL, Vendor_C TEXT, Stk TEXT, Std_LT REAL,
                Shipped_Dt TEXT, ShipTo TEXT, ViaDesc TEXT, Tracking_No TEXT,
                AddntlTracking TEXT, CreditedInv TEXT, Invoice_Line_Memo TEXT,
                Lot_No_Qty TEXT, Manuf_Charge TEXT, Inv_Credit TEXT,
                CM_Reason TEXT, User_Id TEXT, Cust_Item TEXT, Buyer_Name TEXT,
                Issue_Date TEXT, Memo_1 TEXT, Memo_2 TEXT
            )""")
            
            conn.commit()
            print("✅ Tabla invoiced creada/verificada exitosamente")
            
        except Exception as e:
            print(f"❌ Error creando tabla invoiced: {e}")
        finally:
            if 'conn' in locals():
                conn.close()
        
    def get_last_date_in_database(self) -> Optional[str]:
        """
        Obtiene la última fecha de facturación en la base de datos
        Returns: String con fecha en formato YYYY-MM-DD o None si no hay datos
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Verificar si la tabla existe
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='invoiced'")
            if not cursor.fetchone():
                return None
            
            # Obtener la fecha máxima
            cursor.execute("""
                SELECT MAX(Inv_Dt) 
                FROM invoiced 
                WHERE Inv_Dt IS NOT NULL 
                AND Inv_Dt != '' 
                AND Inv_Dt != 'NULL'
            """)
            
            result = cursor.fetchone()
            last_date = result[0] if result and result[0] else None
            
            print(f"🔍 Última fecha en BD: {last_date}")
            return last_date
            
        except Exception as e:
            print(f"❌ Error obteniendo última fecha: {e}")
            return None
        finally:
            if 'conn' in locals():
                conn.close()
    
    def delete_records_from_date(self, from_date: str) -> int:
        """
        Elimina registros desde una fecha específica (inclusive)
        Args:
            from_date: Fecha en formato YYYY-MM-DD
        Returns:
            Número de registros eliminados
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Eliminar registros desde la fecha
            cursor.execute("DELETE FROM invoiced WHERE Inv_Dt >= ?", (from_date,))
            deleted_count = cursor.rowcount
            
            conn.commit()
            print(f"🗑️ Eliminados {deleted_count} registros desde {from_date}")
            
            return deleted_count
            
        except Exception as e:
            print(f"❌ Error eliminando registros: {e}")
            return 0
        finally:
            if 'conn' in locals():
                conn.close()
    
    def get_historical_excel_files(self) -> List[str]:
        """
        Obtiene lista de archivos Excel del directorio histórico
        Returns: Lista de nombres de archivos .xlsx
        """
        files = []
        
        if not os.path.exists(self.historical_dir):
            print(f"⚠️ Directorio histórico no encontrado: {self.historical_dir}")
            return files
        
        try:
            for filename in os.listdir(self.historical_dir):
                if filename.endswith('.xlsx') and not filename.startswith('~$'):
                    files.append(filename)
            
            # Ordenar por año (extraer año del nombre)
            files.sort(key=lambda x: self.extract_year_from_filename(x) or "0000")
            
            print(f"📂 Encontrados {len(files)} archivos históricos")
            return files
            
        except Exception as e:
            print(f"❌ Error listando archivos históricos: {e}")
            return files
    
    def extract_year_from_filename(self, filename: str) -> Optional[str]:
        """Extrae año del nombre del archivo"""
        year_pattern = r'(20[1-3][0-9])'
        match = re.search(year_pattern, filename)
        return match.group(1) if match else None
    
    def read_and_process_excel(self, file_path: str, source_label: str = "") -> Tuple[pd.DataFrame, List[str]]:
        """
        Lee y procesa un archivo Excel de facturas
        Args:
            file_path: Ruta completa del archivo
            source_label: Etiqueta para identificar el origen
        Returns:
            Tuple con DataFrame procesado y lista de errores
        """
        errors = []
        filename = os.path.basename(file_path)
        
        try:
            print(f"📄 Procesando: {filename} {source_label}")
            
            # Verificar que el archivo existe
            if not os.path.exists(file_path):
                errors.append(f"Archivo no encontrado: {file_path}")
                return pd.DataFrame(), errors
            
            # Leer archivo Excel - Detectar tipo por ubicación
            if file_path == self.current_file:
                # Archivo diario (invExp.xlsx) - headers en row 4
                df = pd.read_excel(file_path, skiprows=3, dtype=str)
                print(f"   📋 Archivo DIARIO - usando skiprows=3")
            else:
                # Archivos históricos - headers en row 1
                df = pd.read_excel(file_path, skiprows=0, dtype=str)
                print(f"   📋 Archivo HISTÓRICO - usando skiprows=0")
            
            if df.empty:
                errors.append(f"Archivo vacío: {filename}")
                return pd.DataFrame(), errors
            
            print(f"   📊 Filas iniciales: {len(df)}")
            print(f"   📋 Columnas encontradas: {list(df.columns)[:10]}...")  # Debug - mostrar primeras 10 columnas
            
            # Rellenar valores nulos
            df = df.fillna('')
            
            # Limpiar comillas simples
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].astype(str).str.replace("'", '', regex=False)
            
            # ✅ MAPEO DE COLUMNAS MEJORADO Y UNIFICADO
            column_mapping = {
                # Mapeo estándar para archivos diarios
                'Entity-Code': 'Entity_Code',
                'Invoice-No': 'Invoice_No',
                'SO-No': 'SO_No',
                'Cust Name': 'Cust_Name',
                'PO-No': 'PO_No',
                'A/C': 'AC',
                'Spr-C': 'Spr_C',
                'O-Cd': 'O_Cd',
                'Item Number': 'Item_Number',
                'S-Qty': 'S_Qty',
                'Inv-Dt': 'Inv_Dt',
                'Inv-Time': 'Inv_Time',
                'Ord-Dt': 'Ord_Dt',
                'Due-Dt': 'Due_Dt',
                'TimeToConfirm(hr)': 'TimeToConfirm_hr',
                'Plan-FillRate': 'Plan_FillRate',
                'TAT to fill an order': 'TAT_to_fill_an_order',
                'Cust-FillRate': 'Cust_FillRate',
                'Std Cost': 'Std_Cost',
                'Vendor-C': 'Vendor_C',
                'Std LT': 'Std_LT',
                'Shipped-Dt': 'Shipped_Dt',
                'Tracking No': 'Tracking_No',
                'Invoice Line Memo': 'Invoice_Line_Memo',
                'Lot No (Qty)': 'Lot_No_Qty',
                'Manuf Charge': 'Manuf_Charge',
                'Inv-Credit': 'Inv_Credit',
                'CM-Reason': 'CM_Reason',
                'User-Id': 'User_Id',
                'Cust-Item': 'Cust_Item',
                'Buyer-Name': 'Buyer_Name',
                'Issue-Date': 'Issue_Date',
                'Memo[1]': 'Memo_1',
                'Memo[2]': 'Memo_2',
                
                # ✅ MAPEO ADICIONAL PARA ARCHIVOS HISTÓRICOS (sin guiones)
                'Entity Code': 'Entity_Code',
                'Invoice No': 'Invoice_No',
                'SO No': 'SO_No',
                'PO No': 'PO_No',
                'Spr C': 'Spr_C',
                'O Cd': 'O_Cd',
                'Inv Dt': 'Inv_Dt',
                'Inv Time': 'Inv_Time',
                'Ord Dt': 'Ord_Dt',
                'Due Dt': 'Due_Dt',
                'Plan FillRate': 'Plan_FillRate',
                'Cust FillRate': 'Cust_FillRate',
                'Vendor C': 'Vendor_C',
                'Shipped Dt': 'Shipped_Dt',
                'Inv Credit': 'Inv_Credit',
                'CM Reason': 'CM_Reason',
                'User Id': 'User_Id',
                'Cust Item': 'Cust_Item',
                'Buyer Name': 'Buyer_Name',
                'Issue Date': 'Issue_Date',
                
                # Variaciones adicionales comunes
                'Req-Date': 'Req_Date',
                'Req Date': 'Req_Date',
                'Pln-Ship': 'Pln_Ship',
                'Pln Ship': 'Pln_Ship'
            }
            
            # Renombrar columnas que existan
            df = df.rename(columns=column_mapping)
            
            # ✅ DEFINIR COLUMNAS REQUERIDAS (SIN file_source ni load_timestamp)
            required_columns = [
                'Entity_Code', 'Invoice_No', 'Line', 'SO_No', 'Customer', 'Type', 'Cust_Name',
                'PO_No', 'AC', 'Spr_C', 'O_Cd', 'Proj', 'TC', 'Item_Number', 'Description',
                'UM', 'S_Qty', 'Price', 'Total', 'Curr', 'CustReqDt', 'Req_Date', 'Pln_Ship',
                'Inv_Dt', 'Inv_Time', 'Ord_Dt', 'SODueDt', 'Due_Dt', 'TimeToConfirm_hr',
                'Plan_FillRate', 'TAT_to_fill_an_order', 'Cust_FillRate', 'SH', 'DG', 'Std_Cost',
                'Vendor_C', 'Stk', 'Std_LT', 'Shipped_Dt', 'ShipTo', 'ViaDesc', 'Tracking_No',
                'AddntlTracking', 'CreditedInv', 'Invoice_Line_Memo', 'Lot_No_Qty',
                'Manuf_Charge', 'Inv_Credit', 'CM_Reason', 'User_Id', 'Cust_Item',
                'Buyer_Name', 'Issue_Date', 'Memo_1', 'Memo_2'
            ]
            
            # Agregar columnas faltantes con valores vacíos
            missing_columns = []
            for col in required_columns:
                if col not in df.columns:
                    df[col] = ''
                    missing_columns.append(col)
            
            if missing_columns:
                print(f"   ➕ Agregadas {len(missing_columns)} columnas faltantes")
            
            # ✅ SELECCIONAR SOLO LAS COLUMNAS REQUERIDAS EN EL ORDEN CORRECTO
            df = df[required_columns]
            
            # Validar columnas críticas
            if 'Invoice_No' not in df.columns or 'Inv_Dt' not in df.columns:
                errors.append(f"Columnas críticas faltantes en {filename}")
                return pd.DataFrame(), errors
            
            # Filtrar filas sin número de factura
            initial_count = len(df)
            df = df[df['Invoice_No'].str.strip() != '']
            filtered_count = initial_count - len(df)
            
            if filtered_count > 0:
                print(f"   🔍 Filtradas {filtered_count} filas sin Invoice_No")
            
            # ✅ CONVERTIR COLUMNAS NUMÉRICAS
            numeric_columns = ['S_Qty', 'Price', 'Total', 'TimeToConfirm_hr', 
                             'Plan_FillRate', 'TAT_to_fill_an_order', 'Cust_FillRate', 
                             'Std_Cost', 'Std_LT']
            
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            print(f"   ✅ Procesadas {len(df)} filas válidas")
            
            return df, errors
            
        except Exception as e:
            error_msg = f"Error procesando {filename}: {str(e)}"
            print(f"   ❌ {error_msg}")
            errors.append(error_msg)
            return pd.DataFrame(), errors
    
    def insert_dataframe_to_db(self, df: pd.DataFrame) -> int:
        """
        Inserta DataFrame en la tabla invoiced
        Returns: Número de registros insertados
        """
        if df.empty:
            return 0
        
        try:
            conn = sqlite3.connect(self.db_path)
            
            # ✅ INSERTAR DATOS SIN ESPECIFICAR COLUMNAS EXTRA
            df.to_sql('invoiced', conn, if_exists='append', index=False)
            conn.commit()
            
            inserted_count = len(df)
            print(f"💾 Insertados {inserted_count} registros en BD")
            
            return inserted_count
            
        except Exception as e:
            print(f"❌ Error insertando en BD: {e}")
            return 0
        finally:
            if 'conn' in locals():
                conn.close()
    
    def get_database_stats(self) -> Dict[str, Any]:
        """Obtiene estadísticas de la tabla invoiced"""
        try:
            # ✅ ASEGURAR QUE LA TABLA EXISTE ANTES DE CONSULTAR
            self.ensure_table_exists()
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Verificar si la tabla existe
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='invoiced'")
            if not cursor.fetchone():
                return {'table_exists': False, 'total_records': 0, 'files_loaded': 0}
            
            # Estadísticas básicas
            cursor.execute("SELECT COUNT(*) FROM invoiced")
            total_records = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(DISTINCT Invoice_No) FROM invoiced WHERE Invoice_No IS NOT NULL AND Invoice_No != ''")
            unique_invoices = cursor.fetchone()[0]
            
            cursor.execute("SELECT MIN(Inv_Dt), MAX(Inv_Dt) FROM invoiced WHERE Inv_Dt IS NOT NULL AND Inv_Dt != ''")
            date_range = cursor.fetchone()
            
            # ✅ CALCULAR ARCHIVOS CARGADOS BASADO EN ENTIDADES ÚNICAS
            cursor.execute("SELECT COUNT(DISTINCT Entity_Code) FROM invoiced WHERE Entity_Code IS NOT NULL AND Entity_Code != ''")
            entities_count = cursor.fetchone()[0]
            
            # ✅ ESTIMAR ARCHIVOS CARGADOS BASADO EN RANGO DE AÑOS
            cursor.execute("""
                SELECT COUNT(DISTINCT strftime('%Y', Inv_Dt)) 
                FROM invoiced 
                WHERE Inv_Dt IS NOT NULL AND Inv_Dt != '' AND Inv_Dt != 'NULL'
            """)
            years_with_data = cursor.fetchone()[0] or 0
            
            return {
                'table_exists': True,
                'total_records': total_records,
                'unique_invoices': unique_invoices,
                'min_date': date_range[0] if date_range[0] else None,
                'max_date': date_range[1] if date_range[1] else None,
                'files_loaded': years_with_data,  # ✅ Usar años con datos como proxy
                'entities_count': entities_count
            }
            
        except Exception as e:
            print(f"❌ Error obteniendo estadísticas: {e}")
            return {'table_exists': False, 'error': str(e), 'files_loaded': 0}
        finally:
            if 'conn' in locals():
                conn.close()
    
    def process_incremental_mode(self) -> Dict[str, Any]:
        """
        MODO INCREMENTAL: Procesa solo el archivo actual con control de duplicados
        """
        print("🔄 INICIANDO MODO INCREMENTAL")
        print("=" * 50)
        
        result = {
            'success': False,
            'mode': 'incremental',
            'message': '',
            'stats': {}
        }
        
        try:
            # 1. Obtener última fecha en BD
            last_date_db = self.get_last_date_in_database()
            
            # 2. Leer archivo actual
            df_current, errors = self.read_and_process_excel(self.current_file, "(ACTUAL)")
            
            if errors:
                result['message'] = f"Errores leyendo archivo: {'; '.join(errors)}"
                return result
            
            if df_current.empty:
                result['message'] = "Archivo actual está vacío"
                return result
            
            # 3. Analizar fechas en archivo actual
            df_current['Inv_Dt'] = pd.to_datetime(df_current['Inv_Dt'], errors='coerce')
            df_current = df_current.dropna(subset=['Inv_Dt'])  # Eliminar filas sin fecha válida
            
            if df_current.empty:
                result['message'] = "No hay fechas válidas en el archivo"
                return result
            
            min_date_file = df_current['Inv_Dt'].min().strftime('%Y-%m-%d')
            max_date_file = df_current['Inv_Dt'].max().strftime('%Y-%m-%d')
            
            print(f"📊 Archivo actual - Rango: {min_date_file} a {max_date_file}")
            print(f"📊 Registros en archivo: {len(df_current)}")
            
            # 4. Lógica de control de duplicados
            deleted_count = 0
            if last_date_db:
                # Si la última fecha de BD está en el rango del archivo, eliminar desde esa fecha
                if last_date_db >= min_date_file:
                    print(f"⚠️ Overlap detectado. Eliminando desde {last_date_db}")
                    deleted_count = self.delete_records_from_date(last_date_db)
                else:
                    print(f"✅ No hay overlap. Última BD: {last_date_db}, Mínima archivo: {min_date_file}")
            else:
                print("ℹ️ No hay datos previos en BD")
            
            # 5. Convertir fechas para inserción
            df_current['Inv_Dt'] = df_current['Inv_Dt'].dt.strftime('%Y-%m-%d')
            
            # 6. Insertar datos
            inserted_count = self.insert_dataframe_to_db(df_current)
            
            # 7. Resultado
            result['success'] = inserted_count > 0
            result['message'] = f"Modo incremental completado"
            result['stats'] = {
                'file_processed': os.path.basename(self.current_file),
                'records_in_file': len(df_current),
                'records_deleted': deleted_count,
                'records_inserted': inserted_count,
                'date_range': f"{min_date_file} a {max_date_file}"
            }
            
            print(f"✅ MODO INCREMENTAL COMPLETADO")
            print(f"   📄 Archivo: {os.path.basename(self.current_file)}")
            print(f"   🗑️ Eliminados: {deleted_count}")
            print(f"   💾 Insertados: {inserted_count}")
            
        except Exception as e:
            result['message'] = f"Error en modo incremental: {str(e)}"
            print(f"❌ Error en modo incremental: {e}")
        
        return result

    def process_full_mode(self) -> Dict[str, Any]:
        """
        MODO COMPLETO: Procesa TODOS los archivos históricos + archivo actual
        """
        print("🚀 INICIANDO MODO COMPLETO")
        print("=" * 50)
        
        result = {
            'success': False,
            'mode': 'complete',
            'message': '',
            'stats': {}
        }
        
        try:
            # 1. Limpiar tabla completa
            print("🗑️ Limpiando tabla completa...")
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM invoiced")
            conn.commit()
            conn.close()
            print("✅ Tabla limpiada")
            
            # 2. Obtener archivos históricos
            historical_files = self.get_historical_excel_files()
            
            total_stats = {
                'files_processed': 0,
                'files_with_errors': 0,
                'total_records': 0,
                'historical_records': 0,
                'current_records': 0,
                'error_details': []
            }
            
            # 3. Procesar archivos históricos
            if historical_files:
                print(f"\n📂 Procesando {len(historical_files)} archivos históricos...")
                
                for filename in historical_files:
                    file_path = os.path.join(self.historical_dir, filename)
                    year = self.extract_year_from_filename(filename)
                    
                    df_hist, errors = self.read_and_process_excel(
                        file_path, 
                        f"(HISTÓRICO {year})"
                    )
                    
                    if errors:
                        total_stats['files_with_errors'] += 1
                        total_stats['error_details'].extend(errors)
                        print(f"   ❌ Errores en {filename}: {'; '.join(errors)}")
                        continue
                    
                    if not df_hist.empty:
                        # Convertir fechas para inserción
                        df_hist['Inv_Dt'] = pd.to_datetime(df_hist['Inv_Dt'], errors='coerce')
                        df_hist = df_hist.dropna(subset=['Inv_Dt'])
                        df_hist['Inv_Dt'] = df_hist['Inv_Dt'].dt.strftime('%Y-%m-%d')
                        
                        inserted = self.insert_dataframe_to_db(df_hist)
                        total_stats['files_processed'] += 1
                        total_stats['historical_records'] += inserted
                        total_stats['total_records'] += inserted
                        
                        print(f"   ✅ {filename}: {inserted} registros")
                    else:
                        print(f"   ⚠️ {filename}: archivo vacío")
            else:
                print("⚠️ No se encontraron archivos históricos")
            
            # 4. Procesar archivo actual
            print(f"\n📄 Procesando archivo actual...")
            df_current, current_errors = self.read_and_process_excel(
                self.current_file, 
                "(ACTUAL)"
            )
            
            if current_errors:
                total_stats['error_details'].extend(current_errors)
                print(f"❌ Errores en archivo actual: {'; '.join(current_errors)}")
            else:
                if not df_current.empty:
                    # Convertir fechas para inserción
                    df_current['Inv_Dt'] = pd.to_datetime(df_current['Inv_Dt'], errors='coerce')
                    df_current = df_current.dropna(subset=['Inv_Dt'])
                    df_current['Inv_Dt'] = df_current['Inv_Dt'].dt.strftime('%Y-%m-%d')
                    
                    current_inserted = self.insert_dataframe_to_db(df_current)
                    total_stats['current_records'] = current_inserted
                    total_stats['total_records'] += current_inserted
                    
                    print(f"✅ Archivo actual: {current_inserted} registros")
                else:
                    print("⚠️ Archivo actual está vacío")
            
            # 5. Resultado final
            result['success'] = total_stats['total_records'] > 0
            result['message'] = f"Modo completo finalizado"
            result['stats'] = total_stats
            
            print(f"\n🎉 MODO COMPLETO FINALIZADO")
            print(f"   📂 Archivos históricos procesados: {total_stats['files_processed']}")
            print(f"   📄 Registros históricos: {total_stats['historical_records']:,}")
            print(f"   📄 Registros actuales: {total_stats['current_records']:,}")
            print(f"   📊 TOTAL REGISTROS: {total_stats['total_records']:,}")
            
            if total_stats['files_with_errors'] > 0:
                print(f"   ⚠️ Archivos con errores: {total_stats['files_with_errors']}")
            
        except Exception as e:
            result['message'] = f"Error en modo completo: {str(e)}"
            print(f"❌ Error crítico en modo completo: {e}")
        
        return result
    
    def clear_table(self) -> bool:
        """
        Limpia completamente la tabla invoiced
        Returns: True si fue exitoso
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM invoiced")
            conn.commit()
            conn.close()
            print("🗑️ Tabla invoiced limpiada completamente")
            return True
        except Exception as e:
            print(f"❌ Error limpiando tabla: {e}")
            return False
    
    def validate_files_availability(self) -> Dict[str, Any]:
        """
        Valida la disponibilidad de archivos necesarios
        Returns: Diccionario con estado de archivos
        """
        validation = {
            'current_file_exists': False,
            'current_file_path': self.current_file,
            'historical_dir_exists': False,
            'historical_dir_path': self.historical_dir,
            'historical_files_count': 0,
            'historical_files': [],
            'all_ready': False
        }
        
        try:
            # Validar archivo actual
            validation['current_file_exists'] = os.path.exists(self.current_file)
            
            # Validar directorio histórico
            validation['historical_dir_exists'] = os.path.exists(self.historical_dir)
            
            if validation['historical_dir_exists']:
                historical_files = self.get_historical_excel_files()
                validation['historical_files_count'] = len(historical_files)
                validation['historical_files'] = historical_files
            
            # Determinar si todo está listo
            validation['all_ready'] = (
                validation['current_file_exists'] and 
                validation['historical_dir_exists']
            )
            
            print(f"📋 VALIDACIÓN DE ARCHIVOS:")
            print(f"   📄 Archivo actual: {'✅' if validation['current_file_exists'] else '❌'}")
            print(f"   📂 Directorio histórico: {'✅' if validation['historical_dir_exists'] else '❌'}")
            print(f"   📚 Archivos históricos: {validation['historical_files_count']}")
            
        except Exception as e:
            print(f"❌ Error en validación: {e}")
        
        return validation