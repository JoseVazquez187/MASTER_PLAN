import sqlite3
import pandas as pd
from datetime import datetime
import logging
import os

# Configurar logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DemandValidator:
    def __init__(self, db_path="J:/Departments/Operations/Shared/IT Administration/Python/IRPT/R4Database/R4Database.db"):
        """
        Inicializa el validador de demanda con la ruta de la base de datos
        
        Args:
            db_path (str): Ruta a la base de datos SQLite
        """
        self.db_path = db_path
        self.connection = None
        
    def connect_database(self):
        """Establece conexión con la base de datos"""
        try:
            self.connection = sqlite3.connect(self.db_path)
            logger.info("Conexión a la base de datos establecida exitosamente")
            return True
        except Exception as e:
            logger.error(f"Error al conectar con la base de datos: {e}")
            return False
    
    def close_connection(self):
        """Cierra la conexión con la base de datos"""
        if self.connection:
            self.connection.close()
            logger.info("Conexión a la base de datos cerrada")
    
    def get_expedite_data(self):
        """
        Obtiene los datos de la tabla expedite según el query especificado
        
        Returns:
            pandas.DataFrame: DataFrame con los datos de expedite
        """
        query = """
        SELECT
            expedite.EntityGroup,
            expedite.AC,
            expedite.ItemNo,
            expedite.Description,
            expedite.ReqQty,
            expedite.DemandSource,
            expedite.DemandType,
            expedite.FillDoc,
            REPLACE(expedite.Ref, '-0', '') AS Ref_Limpio,
            expedite.Sub,
            REPLACE(expedite.Ref, '-0', '') || '-' || expedite.Sub AS key,
            expedite.ShipDate
        FROM
            expedite
        WHERE expedite.DemandSource = expedite.ItemNo
        ORDER BY
            SUBSTR(expedite.ShipDate, 7, 2) || '-' ||
            SUBSTR(expedite.ShipDate, 1, 2) || '-' ||
            SUBSTR(expedite.ShipDate, 4, 2)
        ASC;
        """
        
        try:
            df = pd.read_sql_query(query, self.connection)
            logger.info(f"Se obtuvieron {len(df)} registros de la tabla expedite")
            return df
        except Exception as e:
            logger.error(f"Error al obtener datos de expedite: {e}")
            return pd.DataFrame()
    
    def get_fcst_data(self):
        """
        Obtiene los datos de la tabla fcst con OpenQty > 0
        
        Returns:
            pandas.DataFrame: DataFrame con los datos de fcst
        """
        query = """
        SELECT 
            Entity,
            Proj,
            AC,
            ConfigID,
            FcstNo,
            Description,
            ItemNo,
            Rev,
            UM,
            PlannedBy,
            ReqDate,
            QtyFcst,
            OpenQty,
            WO
        FROM fcst 
        WHERE fcst.OpenQty > 0
        """
        
        try:
            df = pd.read_sql_query(query, self.connection)
            logger.info(f"Se obtuvieron {len(df)} registros de la tabla fcst")
            return df
        except Exception as e:
            logger.error(f"Error al obtener datos de fcst: {e}")
            return pd.DataFrame()
    
    def get_wo_data(self):
        """
        Obtiene los datos de la tabla WorkOrder (WOInquiry)
        
        Returns:
            pandas.DataFrame: DataFrame con los datos de WorkOrder
        """
        query = """
        SELECT 
            WOInquiry.ItemNumber,
            WOInquiry.Description,
            WOInquiry.Srt,
            WOInquiry.AC,
            WOInquiry.CreateDt,
            WOInquiry.OpnQ,
            WOInquiry.SO_FCST,
            WOInquiry.Sub,
            WOInquiry.WONo,
            WOInquiry.Prtdate
        FROM WOInquiry 
        WHERE WOInquiry.Srt LIKE 'G%' OR WOInquiry.Srt LIKE 'K%' 
        ORDER BY WOInquiry.CreateDt ASC;
        """
        
        try:
            df = pd.read_sql_query(query, self.connection)
            logger.info(f"Se obtuvieron {len(df)} registros de la tabla WOInquiry")
            return df
        except Exception as e:
            logger.error(f"Error al obtener datos de WOInquiry: {e}")
            return pd.DataFrame()
    
    def get_so_data(self):
        """
        Obtiene los datos de la tabla sales_order_table
        
        Returns:
            pandas.DataFrame: DataFrame con los datos de sales_order_table
        """
        query = """
        SELECT 
            sales_order_table.Entity,
            sales_order_table.Proj,
            sales_order_table.SO_No,
            sales_order_table.Ln,
            sales_order_table.Ord_Cd,
            sales_order_table.Spr_CD,
            sales_order_table.Order_Dt,
            sales_order_table.Req_Dt,
            sales_order_table.Prd_Dt,
            sales_order_table.Cust_PO,
            sales_order_table.AC,
            sales_order_table.Item_Number,
            sales_order_table.Description,
            sales_order_table.UM,
            sales_order_table.ML,
            sales_order_table.Opn_Q,
            sales_order_table.Issue_Q,
            sales_order_table.WO_No,
            sales_order_table.WO_Due_dt,
            sales_order_table.SO_No|| '-' ||sales_order_table.Ln AS KEY
        FROM sales_order_table;
        """
        
        try:
            df = pd.read_sql_query(query, self.connection)
            logger.info(f"Se obtuvieron {len(df)} registros de la tabla sales_order_table")
            return df
        except Exception as e:
            logger.error(f"Error al obtener datos de sales_order_table: {e}")
            return pd.DataFrame()
        """
        Obtiene los datos de la tabla WorkOrder (WOInquiry)
        
        Returns:
            pandas.DataFrame: DataFrame con los datos de WorkOrder
        """
        query = """
        SELECT 
            WOInquiry.ItemNumber,
            WOInquiry.Description,
            WOInquiry.Srt,
            WOInquiry.AC,
            WOInquiry.CreateDt,
            WOInquiry.OpnQ,
            WOInquiry.SO_FCST,
            WOInquiry.Sub,
            WOInquiry.WONo,
            WOInquiry.Prtdate
        FROM WOInquiry 
        WHERE WOInquiry.Srt LIKE 'G%' OR WOInquiry.Srt LIKE 'K%' 
        ORDER BY WOInquiry.CreateDt ASC;
        """
        
        try:
            df = pd.read_sql_query(query, self.connection)
            logger.info(f"Se obtuvieron {len(df)} registros de la tabla WOInquiry")
            return df
        except Exception as e:
            logger.error(f"Error al obtener datos de WOInquiry: {e}")
            return pd.DataFrame()
    
    def validate_fcst_demand(self, expedite_row, fcst_df):
        """
        Valida demanda tipo FCST usando Ref_Limpio y comparando cantidades
        
        Args:
            expedite_row: Fila de expedite a validar
            fcst_df: DataFrame con datos de fcst
            
        Returns:
            dict: Resultado de la validación
        """
        ref_limpio = str(expedite_row['Ref_Limpio']).strip()
        
        # Convertir ReqQty a numérico, manejar posibles errores
        try:
            req_qty = float(expedite_row['ReqQty']) if expedite_row['ReqQty'] != '' else 0
        except (ValueError, TypeError):
            req_qty = 0
        
        # Debug: Log para verificar la búsqueda
        logger.debug(f"Buscando FCST: Ref_Limpio='{ref_limpio}' en columna FcstNo")
        
        # Convertir FcstNo a string para comparación segura
        fcst_df_copy = fcst_df.copy()
        fcst_df_copy['FcstNo'] = fcst_df_copy['FcstNo'].astype(str).str.strip()
        
        # Buscar en fcst por FcstNo usando Ref_Limpio
        matching_fcst = fcst_df_copy[fcst_df_copy['FcstNo'] == ref_limpio]
        
        # Debug adicional
        if matching_fcst.empty:
            # Verificar si existe con otros valores similares
            similar_matches = fcst_df_copy[fcst_df_copy['FcstNo'].str.contains(ref_limpio, na=False)]
            logger.debug(f"No encontrado exacto para '{ref_limpio}'. Similares encontrados: {len(similar_matches)}")
            if len(similar_matches) > 0:
                logger.debug(f"Valores similares: {similar_matches['FcstNo'].head().tolist()}")
        
        if not matching_fcst.empty:
            # Tomar el primer match (debería ser único)
            fcst_record = matching_fcst.iloc[0]
            
            # Convertir OpenQty a numérico, manejar posibles errores
            try:
                fcst_open_qty = float(fcst_record['OpenQty']) if fcst_record['OpenQty'] != '' else 0
            except (ValueError, TypeError):
                fcst_open_qty = 0
            
            # Validar cantidades
            qty_status = "OK" if fcst_open_qty >= req_qty else "INSUFFICIENT"
            qty_difference = fcst_open_qty - req_qty
            
            logger.debug(f"FCST encontrado: {ref_limpio}, OpenQty={fcst_open_qty}, ReqQty={req_qty}")
            
            return {
                'status': 'FOUND',
                'source_table': 'fcst',
                'matches_found': len(matching_fcst),
                'fcst_open_qty': fcst_open_qty,
                'req_qty': req_qty,
                'qty_difference': qty_difference,
                'qty_status': qty_status,
                'fcst_number': ref_limpio,
                'details': matching_fcst.to_dict('records')
            }
        else:
            logger.debug(f"FCST NO encontrado para: {ref_limpio}")
            return {
                'status': 'NOT_FOUND',
                'source_table': 'fcst',
                'matches_found': 0,
                'fcst_number': ref_limpio,
                'req_qty': req_qty,
                'details': None
            }
    
    def validate_wom_demand(self, expedite_row, wo_df):
        """
        Valida demanda tipo WOM
        
        Args:
            expedite_row: Fila de expedite a validar
            wo_df: DataFrame con datos de WorkOrder
            
        Returns:
            dict: Resultado de la validación
        """
        # Buscar en WorkOrder por WONo usando Ref_Limpio
        ref_limpio = expedite_row['Ref_Limpio']
        matching_wo = wo_df[wo_df['WONo'] == ref_limpio]
        
        if not matching_wo.empty:
            return {
                'status': 'FOUND',
                'source_table': 'WorkOrder',
                'matches_found': len(matching_wo),
                'details': matching_wo.to_dict('records')
            }
        else:
            return {
                'status': 'NOT_FOUND',
                'source_table': 'WorkOrder',
                'matches_found': 0,
                'details': None
            }
    
    def validate_so_demand(self, expedite_row, so_df):
        """
        Valida demanda tipo SO usando la key y comparando cantidades
        
        Args:
            expedite_row: Fila de expedite a validar
            so_df: DataFrame con datos de sales_order_table
            
        Returns:
            dict: Resultado de la validación
        """
        expedite_key = str(expedite_row['key']).strip()
        
        # Convertir ReqQty a numérico, manejar posibles errores
        try:
            req_qty = float(expedite_row['ReqQty']) if expedite_row['ReqQty'] != '' else 0
        except (ValueError, TypeError):
            req_qty = 0
        
        # Debug: Log para verificar la búsqueda
        logger.debug(f"Buscando SO: key='{expedite_key}' en columna KEY")
        
        # Convertir KEY a string para comparación segura
        so_df_copy = so_df.copy()
        so_df_copy['KEY'] = so_df_copy['KEY'].astype(str).str.strip()
        
        # Buscar en sales_order_table por KEY usando expedite.key
        matching_so = so_df_copy[so_df_copy['KEY'] == expedite_key]
        
        # Debug adicional
        if matching_so.empty:
            # Verificar si existe con otros valores similares
            similar_matches = so_df_copy[so_df_copy['KEY'].str.contains(expedite_key, na=False)]
            logger.debug(f"No encontrado exacto para '{expedite_key}'. Similares encontrados: {len(similar_matches)}")
            if len(similar_matches) > 0:
                logger.debug(f"Valores similares: {similar_matches['KEY'].head().tolist()}")
        
        if not matching_so.empty:
            # Tomar el primer match (debería ser único)
            so_record = matching_so.iloc[0]
            
            # Convertir Opn_Q a numérico, manejar posibles errores
            try:
                so_open_qty = float(so_record['Opn_Q']) if so_record['Opn_Q'] != '' else 0
            except (ValueError, TypeError):
                so_open_qty = 0
            
            # Validar cantidades
            qty_status = "OK" if so_open_qty >= req_qty else "INSUFFICIENT"
            qty_difference = so_open_qty - req_qty
            
            logger.debug(f"SO encontrado: {expedite_key}, Opn_Q={so_open_qty}, ReqQty={req_qty}")
            
            return {
                'status': 'FOUND',
                'source_table': 'sales_order_table',
                'matches_found': len(matching_so),
                'so_open_qty': so_open_qty,
                'req_qty': req_qty,
                'qty_difference': qty_difference,
                'qty_status': qty_status,
                'so_number': so_record['SO_No'],
                'so_line': so_record['Ln'],
                'so_key': expedite_key,
                'wo_number': so_record['WO_No'],
                'customer_po': so_record['Cust_PO'],
                'details': matching_so.to_dict('records')
            }
        else:
            logger.debug(f"SO NO encontrado para: {expedite_key}")
            return {
                'status': 'NOT_FOUND',
                'source_table': 'sales_order_table',
                'matches_found': 0,
                'so_key': expedite_key,
                'req_qty': req_qty,
                'details': None
            }
    
    def validate_safe_demand(self, expedite_row):
        """
        Valida demanda tipo SAFE (placeholder - necesitas especificar la tabla y lógica)
        
        Args:
            expedite_row: Fila de expedite a validar
            
        Returns:
            dict: Resultado de la validación
        """
        # TODO: Implementar validación para SAFE cuando especifiques la tabla/lógica
        return {
            'status': 'PENDING_IMPLEMENTATION',
            'source_table': 'SAFE',
            'message': 'Validación SAFE pendiente de implementación'
        }
    
    def validate_single_demand(self, expedite_row, fcst_df, wo_df, so_df):
        """
        Valida una sola demanda según su tipo
        
        Args:
            expedite_row: Fila de expedite a validar
            fcst_df: DataFrame con datos de fcst
            wo_df: DataFrame con datos de WorkOrder
            so_df: DataFrame con datos de sales_order_table
            
        Returns:
            dict: Resultado de la validación
        """
        try:
            demand_type = expedite_row['DemandType']
            
            validation_result = {
                'expedite_key': expedite_row['key'],
                'item_no': expedite_row['ItemNo'],
                'demand_type': demand_type,
                'req_qty': expedite_row['ReqQty'],
                'ship_date': expedite_row['ShipDate']
            }
            
            if demand_type == 'FCST':
                validation_result.update(self.validate_fcst_demand(expedite_row, fcst_df))
            elif demand_type == 'WOM':
                validation_result.update(self.validate_wom_demand(expedite_row, wo_df))
            elif demand_type == 'SO':
                validation_result.update(self.validate_so_demand(expedite_row, so_df))
            elif demand_type == 'SAFE':
                validation_result.update(self.validate_safe_demand(expedite_row))
            else:
                validation_result.update({
                    'status': 'UNKNOWN_DEMAND_TYPE',
                    'message': f'Tipo de demanda desconocido: {demand_type}'
                })
            
            return validation_result
            
        except Exception as e:
            logger.error(f"Error validando demanda {expedite_row.get('key', 'UNKNOWN')}: {e}")
            return {
                'expedite_key': expedite_row.get('key', 'UNKNOWN'),
                'item_no': expedite_row.get('ItemNo', 'UNKNOWN'),
                'demand_type': expedite_row.get('DemandType', 'UNKNOWN'),
                'req_qty': expedite_row.get('ReqQty', 'UNKNOWN'),
                'ship_date': expedite_row.get('ShipDate', 'UNKNOWN'),
                'status': 'VALIDATION_ERROR',
                'message': f'Error durante validación: {str(e)}'
            }
    
    def validate_all_demands_efficient(self, export_excel=True):
        """
        Ejecuta la validación completa usando merges eficientes
        
        Args:
            export_excel (bool): Si True, exporta resultados a Excel
        
        Returns:
            tuple: (results_df, excel_filename) o (results_df, None)
        """
        if not self.connect_database():
            return pd.DataFrame(), None
        
        try:
            # Obtener todos los datos necesarios
            logger.info("Obteniendo datos de expedite...")
            expedite_df = self.get_expedite_data()
            
            logger.info("Obteniendo datos de fcst...")
            fcst_df = self.get_fcst_data()
            
            logger.info("Obteniendo datos de WorkOrder...")
            wo_df = self.get_wo_data()
            
            logger.info("Obteniendo datos de SalesOrder...")
            so_df = self.get_so_data()
            
            if expedite_df.empty:
                logger.warning("No se encontraron datos en expedite")
                return pd.DataFrame(), None
            
            logger.info(f"Iniciando validación eficiente de {len(expedite_df)} demandas...")
            
            # Preparar DataFrames para merge
            logger.info("Preparando datos para merge...")
            
            # Limpiar y preparar datos FCST
            if not fcst_df.empty:
                fcst_df = fcst_df.copy()
                fcst_df['FcstNo'] = fcst_df['FcstNo'].astype(str).str.strip()
                fcst_df['OpenQty'] = pd.to_numeric(fcst_df['OpenQty'], errors='coerce').fillna(0)
                # Agregar prefijo para evitar conflictos de columnas
                fcst_df = fcst_df.add_prefix('fcst_')
                fcst_df.rename(columns={'fcst_FcstNo': 'Ref_Limpio'}, inplace=True)
            
            # Limpiar y preparar datos WO
            if not wo_df.empty:
                wo_df = wo_df.copy()
                logger.debug(f"Columnas originales en WO: {wo_df.columns.tolist()}")
                
                # Las columnas correctas según la imagen son: WONo, OpnQ
                if 'WONo' in wo_df.columns:
                    wo_df['WONo'] = wo_df['WONo'].astype(str).str.strip()
                    # Usar WONo directamente como clave de merge
                    wo_df['merge_key'] = wo_df['WONo']
                else:
                    logger.warning("Columna 'WONo' no encontrada en wo_df")
                    wo_df['merge_key'] = ''
                
                if 'OpnQ' in wo_df.columns:
                    wo_df['OpnQ'] = pd.to_numeric(wo_df['OpnQ'], errors='coerce').fillna(0)
                else:
                    logger.warning("Columna 'OpnQ' no encontrada en wo_df")
                    wo_df['OpnQ'] = 0
                
                # Agregar prefijo para evitar conflictos de columnas
                wo_df = wo_df.add_prefix('wo_')
                # Renombrar la clave de merge
                wo_df.rename(columns={'wo_merge_key': 'Ref_Limpio'}, inplace=True)
                
                logger.debug(f"Columnas después de preparar WO: {wo_df.columns.tolist()}")
                logger.debug(f"Muestra de Ref_Limpio en WO: {wo_df['Ref_Limpio'].head().tolist()}")
            else:
                logger.warning("wo_df está vacío")
            
            # Limpiar y preparar datos SO
            if not so_df.empty:
                so_df = so_df.copy()
                so_df['KEY'] = so_df['KEY'].astype(str).str.strip()
                so_df['Opn_Q'] = pd.to_numeric(so_df['Opn_Q'], errors='coerce').fillna(0)
                # Agregar prefijo para evitar conflictos de columnas
                so_df = so_df.add_prefix('so_')
                so_df.rename(columns={'so_KEY': 'key'}, inplace=True)
            
            # Preparar expedite
            expedite_df = expedite_df.copy()
            # Convertir ReqQty a numérico ANTES de cualquier procesamiento
            expedite_df['ReqQty'] = pd.to_numeric(expedite_df['ReqQty'], errors='coerce').fillna(0)
            expedite_df['Ref_Limpio'] = expedite_df['Ref_Limpio'].astype(str).str.strip()
            expedite_df['key'] = expedite_df['key'].astype(str).str.strip()
            
            # Debug: Verificar que ReqQty sea numérico
            logger.info(f"Tipo de ReqQty después de conversión: {expedite_df['ReqQty'].dtype}")
            logger.info(f"Muestra de ReqQty: {expedite_df['ReqQty'].head().tolist()}")
            logger.info(f"Suma total de ReqQty: {expedite_df['ReqQty'].sum()}")
            
            # Realizar merges por tipo de demanda
            logger.info("Realizando merges por tipo de demanda...")
            
            # Separar por tipo de demanda
            fcst_demands = expedite_df[expedite_df['DemandType'] == 'FCST'].copy()
            wom_demands = expedite_df[expedite_df['DemandType'] == 'WOM'].copy()
            so_demands = expedite_df[expedite_df['DemandType'] == 'SO'].copy()
            safe_demands = expedite_df[expedite_df['DemandType'] == 'SAFE'].copy()
            other_demands = expedite_df[~expedite_df['DemandType'].isin(['FCST', 'WOM', 'SO', 'SAFE'])].copy()
            
            results_list = []
            
            # Merge FCST
            if not fcst_demands.empty and not fcst_df.empty:
                logger.info(f"Procesando {len(fcst_demands)} demandas FCST...")
                fcst_merged = fcst_demands.merge(fcst_df, on='Ref_Limpio', how='left')
                fcst_merged = self.process_fcst_merge(fcst_merged)
                results_list.append(fcst_merged)
            elif not fcst_demands.empty:
                fcst_demands = self.add_not_found_columns(fcst_demands, 'FCST')
                results_list.append(fcst_demands)
            
            # Merge WOM
            if not wom_demands.empty and not wo_df.empty:
                logger.info(f"Procesando {len(wom_demands)} demandas WOM...")
                logger.debug(f"Columnas en wo_df: {wo_df.columns.tolist()}")
                logger.debug(f"Columnas en wom_demands: {wom_demands.columns.tolist()}")
                
                wom_merged = wom_demands.merge(wo_df, on='Ref_Limpio', how='left')
                logger.debug(f"Columnas después del merge WOM: {wom_merged.columns.tolist()}")
                
                wom_merged = self.process_wom_merge(wom_merged)
                results_list.append(wom_merged)
            elif not wom_demands.empty:
                logger.info(f"No hay datos de WO, marcando {len(wom_demands)} demandas WOM como NOT_FOUND")
                wom_demands = self.add_not_found_columns(wom_demands, 'WOM')
                results_list.append(wom_demands)
            
            # Merge SO
            if not so_demands.empty and not so_df.empty:
                logger.info(f"Procesando {len(so_demands)} demandas SO...")
                so_merged = so_demands.merge(so_df, on='key', how='left')
                so_merged = self.process_so_merge(so_merged)
                results_list.append(so_merged)
            elif not so_demands.empty:
                so_demands = self.add_not_found_columns(so_demands, 'SO')
                results_list.append(so_demands)
            
            # Procesar SAFE y otros
            if not safe_demands.empty:
                safe_demands = self.add_not_found_columns(safe_demands, 'SAFE', 'PENDING_IMPLEMENTATION')
                results_list.append(safe_demands)
            
            if not other_demands.empty:
                other_demands = self.add_not_found_columns(other_demands, 'UNKNOWN', 'UNKNOWN_DEMAND_TYPE')
                results_list.append(other_demands)
            
            # Combinar todos los resultados
            if results_list:
                logger.info("Combinando resultados...")
                final_results = pd.concat(results_list, ignore_index=True)
                
                # Ordenar por el orden original de expedite
                final_results = final_results.sort_values('ShipDate')
                
                logger.info(f"Validación completada eficientemente. {len(final_results)} demandas procesadas")
            else:
                final_results = pd.DataFrame()
            
            # Exportar a Excel si se solicita
            excel_filename = None
            if export_excel and not final_results.empty:
                logger.info("Exportando resultados a Excel...")
                excel_filename = self.export_efficient_results_to_excel(final_results)
            
            return final_results, excel_filename
            
        except Exception as e:
            logger.error(f"Error durante la validación eficiente: {e}")
            return pd.DataFrame(), None
        finally:
            self.close_connection()
    
    def process_fcst_merge(self, merged_df):
        """Procesa el resultado del merge de FCST"""
        merged_df['Validation_Status'] = merged_df['fcst_OpenQty'].apply(
            lambda x: 'FOUND' if pd.notna(x) else 'NOT_FOUND'
        )
        merged_df['Source_Table'] = 'fcst'
        merged_df['FCST_Open_Qty'] = merged_df['fcst_OpenQty'].fillna(0)
        merged_df['Qty_Difference'] = merged_df['FCST_Open_Qty'] - merged_df['ReqQty']
        merged_df['Qty_Status'] = merged_df.apply(
            lambda row: 'OK' if row['FCST_Open_Qty'] >= row['ReqQty'] and row['Validation_Status'] == 'FOUND' 
            else 'INSUFFICIENT' if row['Validation_Status'] == 'FOUND' 
            else '', axis=1
        )
        merged_df['FCST_Number_Found'] = merged_df.apply(
            lambda row: row['Ref_Limpio'] if row['Validation_Status'] == 'FOUND' else '', axis=1
        )
        return merged_df
    
    def process_wom_merge(self, merged_df):
        """Procesa el resultado del merge de WOM"""
        # Verificar si las columnas existen antes de usarlas
        if 'wo_WONo' in merged_df.columns:
            merged_df['Validation_Status'] = merged_df['wo_WONo'].apply(
                lambda x: 'FOUND' if pd.notna(x) and str(x).strip() != '' else 'NOT_FOUND'
            )
            merged_df['WO_Number_Found'] = merged_df['wo_WONo'].fillna('')
        else:
            # Si no hay columna wo_WONo, significa que no hubo merge exitoso
            merged_df['Validation_Status'] = 'NOT_FOUND'
            merged_df['WO_Number_Found'] = ''
        
        merged_df['Source_Table'] = 'WorkOrder'
        
        # Manejar wo_OpnQ de manera segura
        if 'wo_OpnQ' in merged_df.columns:
            merged_df['WO_Open_Qty'] = pd.to_numeric(merged_df['wo_OpnQ'], errors='coerce').fillna(0)
        else:
            merged_df['WO_Open_Qty'] = 0
            
        merged_df['Qty_Difference'] = merged_df['WO_Open_Qty'] - merged_df['ReqQty']
        merged_df['Qty_Status'] = merged_df.apply(
            lambda row: 'OK' if row['WO_Open_Qty'] >= row['ReqQty'] and row['Validation_Status'] == 'FOUND' 
            else 'INSUFFICIENT' if row['Validation_Status'] == 'FOUND' 
            else '', axis=1
        )
        
        # Agregar información adicional si está disponible
        if 'wo_ItemNumber' in merged_df.columns:
            merged_df['WO_ItemNumber'] = merged_df['wo_ItemNumber'].fillna('')
        if 'wo_AC' in merged_df.columns:
            merged_df['WO_AC_Found'] = merged_df['wo_AC'].fillna('')
        
        return merged_df
    
    def process_so_merge(self, merged_df):
        """Procesa el resultado del merge de SO"""
        merged_df['Validation_Status'] = merged_df['so_SO_No'].apply(
            lambda x: 'FOUND' if pd.notna(x) else 'NOT_FOUND'
        )
        merged_df['Source_Table'] = 'sales_order_table'
        merged_df['SO_Open_Qty'] = merged_df['so_Opn_Q'].fillna(0)
        merged_df['Qty_Difference'] = merged_df['SO_Open_Qty'] - merged_df['ReqQty']
        merged_df['Qty_Status'] = merged_df.apply(
            lambda row: 'OK' if row['SO_Open_Qty'] >= row['ReqQty'] and row['Validation_Status'] == 'FOUND' 
            else 'INSUFFICIENT' if row['Validation_Status'] == 'FOUND' 
            else '', axis=1
        )
        merged_df['SO_Number_Found'] = merged_df['so_SO_No'].fillna('')
        merged_df['SO_Line_Found'] = merged_df['so_Ln'].fillna('')
        merged_df['Customer_PO_Found'] = merged_df['so_Cust_PO'].fillna('')
        merged_df['WO_From_SO'] = merged_df['so_WO_No'].fillna('')
        return merged_df
    
    def add_not_found_columns(self, df, demand_type, status='NOT_FOUND'):
        """Agrega columnas de validación para demandas no encontradas"""
        df = df.copy()
        df['Validation_Status'] = status
        df['Source_Table'] = demand_type
        df['FCST_Open_Qty'] = ''
        df['SO_Open_Qty'] = ''
        df['WO_Open_Qty'] = ''
        df['Qty_Difference'] = ''
        df['Qty_Status'] = ''
        df['FCST_Number_Found'] = ''
        df['SO_Number_Found'] = ''
        df['SO_Line_Found'] = ''
        df['Customer_PO_Found'] = ''
        df['WO_Number_Found'] = ''
        df['WO_From_SO'] = ''
        df['WO_ItemNumber'] = ''
        df['WO_AC_Found'] = ''
        return df
    
    def export_efficient_results_to_excel(self, results_df, filename=None):
        """Exporta los resultados eficientes a Excel"""
        if filename is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'Demand_Validation_Results_Efficient_{timestamp}.xlsx'
        
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                
                # Hoja 1: Resultados principales
                main_columns = [
                    'EntityGroup', 'AC', 'ItemNo', 'Description', 'ReqQty', 'DemandSource', 
                    'DemandType', 'FillDoc', 'Ref_Limpio', 'Sub', 'key', 'ShipDate',
                    'Validation_Status', 'Source_Table', 'Qty_Status', 'Qty_Difference',
                    'FCST_Open_Qty', 'FCST_Number_Found', 
                    'SO_Open_Qty', 'SO_Number_Found', 'SO_Line_Found', 'Customer_PO_Found',
                    'WO_Open_Qty', 'WO_Number_Found', 'WO_From_SO'
                ]
                
                # Seleccionar solo columnas que existen
                available_columns = [col for col in main_columns if col in results_df.columns]
                results_df[available_columns].to_excel(writer, sheet_name='Validation_Results', index=False)
                
                # Hoja 2: Resumen por status
                status_summary = results_df.groupby(['DemandType', 'Validation_Status']).size().reset_index(name='Count')
                status_summary.to_excel(writer, sheet_name='Status_Summary', index=False)
                
                # Hoja 3: Resumen por AC
                ac_summary = results_df.groupby(['AC', 'DemandType', 'Validation_Status']).agg({
                    'ReqQty': 'sum',  # Sumar las cantidades
                    'key': 'count'    # Contar los registros
                }).reset_index()
                ac_summary.columns = ['AC', 'DemandType', 'Validation_Status', 'Total_ReqQty', 'Count']
                # Asegurar que Total_ReqQty sea numérico
                ac_summary['Total_ReqQty'] = pd.to_numeric(ac_summary['Total_ReqQty'], errors='coerce').fillna(0)
                ac_summary.to_excel(writer, sheet_name='AC_Summary', index=False)
                
                # Hoja 4: Solo no encontrados
                not_found = results_df[results_df['Validation_Status'] == 'NOT_FOUND']
                if not not_found.empty:
                    not_found[available_columns].to_excel(writer, sheet_name='Not_Found', index=False)
                
                # Hoja 6: Resumen total por AC (más limpio)
                ac_total_summary = results_df.groupby('AC').agg({
                    'ReqQty': 'sum',
                    'key': 'count'
                }).reset_index()
                ac_total_summary.columns = ['AC', 'Total_ReqQty', 'Total_Demands']
                
                # Agregar columna de demandas encontradas
                found_by_ac = results_df[results_df['Validation_Status'] == 'FOUND'].groupby('AC').agg({
                    'ReqQty': 'sum',
                    'key': 'count'
                }).reset_index()
                found_by_ac.columns = ['AC', 'Found_ReqQty', 'Found_Demands']
                
                # Merge para combinar
                ac_total_summary = ac_total_summary.merge(found_by_ac, on='AC', how='left')
                ac_total_summary['Found_ReqQty'] = ac_total_summary['Found_ReqQty'].fillna(0)
                ac_total_summary['Found_Demands'] = ac_total_summary['Found_Demands'].fillna(0)
                
                # Calcular tasas de éxito
                ac_total_summary['Success_Rate_Demands'] = (ac_total_summary['Found_Demands'] / ac_total_summary['Total_Demands'] * 100).round(2)
                ac_total_summary['Success_Rate_Qty'] = (ac_total_summary['Found_ReqQty'] / ac_total_summary['Total_ReqQty'] * 100).round(2)
                
                # Asegurar que las cantidades sean numéricas
                numeric_cols = ['Total_ReqQty', 'Found_ReqQty', 'Total_Demands', 'Found_Demands']
                for col in numeric_cols:
                    ac_total_summary[col] = pd.to_numeric(ac_total_summary[col], errors='coerce').fillna(0)
                
                ac_total_summary.to_excel(writer, sheet_name='AC_Total_Summary', index=False)
                
                # Hoja 7: Solo encontrados
                found = results_df[results_df['Validation_Status'] == 'FOUND']
                if not found.empty:
                    found[available_columns].to_excel(writer, sheet_name='Found', index=False)
            
            logger.info(f"Archivo Excel eficiente creado: {filename}")
            return filename
            
        except Exception as e:
            logger.error(f"Error al crear archivo Excel eficiente: {e}")
            return None
        """
        Ejecuta la validación completa de todas las demandas
        
        Args:
            export_excel (bool): Si True, exporta resultados a Excel
        
        Returns:
            tuple: (results, excel_filename) o (results, None)
        """
        if not self.connect_database():
            return [], None
        
        try:
            # Obtener todos los datos necesarios
            logger.info("Obteniendo datos de expedite...")
            expedite_df = self.get_expedite_data()
            
            logger.info("Obteniendo datos de fcst...")
            fcst_df = self.get_fcst_data()
            
            logger.info("Obteniendo datos de WorkOrder...")
            wo_df = self.get_wo_data()
            
            logger.info("Obteniendo datos de SalesOrder...")
            so_df = self.get_so_data()
            
            if expedite_df.empty:
                logger.warning("No se encontraron datos en expedite")
                return [], None
            
            # Validar cada demanda
            results = []
            logger.info(f"Iniciando validación de {len(expedite_df)} demandas...")
            
            for index, row in expedite_df.iterrows():
                result = self.validate_single_demand(row, fcst_df, wo_df, so_df)
                results.append(result)
                
                if (index + 1) % 100 == 0:
                    logger.info(f"Procesadas {index + 1} demandas...")
            
            logger.info(f"Validación completada. {len(results)} demandas procesadas")
            
            # Exportar a Excel si se solicita
            excel_filename = None
            if export_excel and results:
                logger.info("Creando matriz de validación...")
                matrix_df = self.create_validation_matrix(expedite_df, results)
                
                logger.info("Exportando resultados a Excel...")
                excel_filename = self.export_to_excel(matrix_df, results)
            
            return results, excel_filename
            
        except Exception as e:
            logger.error(f"Error durante la validación: {e}")
            return [], None
        finally:
            self.close_connection()
    
    def create_validation_matrix(self, expedite_df, results):
        """
        Crea una matriz completa combinando expedite con los resultados de validación
        
        Args:
            expedite_df: DataFrame original de expedite
            results: Lista de resultados de validación
            
        Returns:
            pandas.DataFrame: DataFrame con la matriz completa
        """
        # Convertir resultados a DataFrame para facilitar el merge
        results_df = pd.DataFrame(results)
        
        # Preparar DataFrame base con todos los datos de expedite
        matrix_df = expedite_df.copy()
        
        # Agregar columnas de validación
        if not results_df.empty:
            # Crear diccionario de resultados indexado por key para mapeo rápido
            results_dict = {}
            for result in results:
                key = result.get('expedite_key')
                results_dict[key] = result
            
            # Agregar columnas de validación
            validation_columns = {
                'Validation_Status': [],
                'Source_Table': [],
                'Matches_Found': [],
                'FCST_Open_Qty': [],
                'SO_Open_Qty': [],
                'Qty_Difference': [],
                'Qty_Status': [],
                'Validation_Message': [],
                'Found_Details_Count': [],
                'WO_Numbers_Found': [],
                'FCST_Number_Found': [],
                'SO_Number_Found': [],
                'SO_Line_Found': [],
                'Customer_PO_Found': []
            }
            
            for _, row in matrix_df.iterrows():
                key = row['key']
                result = results_dict.get(key, {})
                
                # Status de validación
                validation_columns['Validation_Status'].append(result.get('status', 'NO_RESULT'))
                validation_columns['Source_Table'].append(result.get('source_table', ''))
                validation_columns['Matches_Found'].append(result.get('matches_found', 0))
                
                # Información específica de FCST
                validation_columns['FCST_Open_Qty'].append(result.get('fcst_open_qty', ''))
                validation_columns['FCST_Number_Found'].append(result.get('fcst_number', ''))
                
                # Información específica de SO
                validation_columns['SO_Open_Qty'].append(result.get('so_open_qty', ''))
                validation_columns['SO_Number_Found'].append(result.get('so_number', ''))
                validation_columns['SO_Line_Found'].append(result.get('so_line', ''))
                validation_columns['Customer_PO_Found'].append(result.get('customer_po', ''))
                
                # Información común
                validation_columns['Qty_Difference'].append(result.get('qty_difference', ''))
                validation_columns['Qty_Status'].append(result.get('qty_status', ''))
                validation_columns['Validation_Message'].append(result.get('message', ''))
                
                # Detalles específicos
                details = result.get('details', [])
                if details and isinstance(details, list):
                    validation_columns['Found_Details_Count'].append(len(details))
                    
                    # Extraer WO Numbers si existen
                    wo_numbers = []
                    for detail in details:
                        if 'WONo' in detail:
                            wo_numbers.append(str(detail['WONo']))
                        elif 'WO_No' in detail:  # Para SO puede ser WO_No
                            wo_numbers.append(str(detail['WO_No']))
                    
                    validation_columns['WO_Numbers_Found'].append('; '.join(wo_numbers))
                else:
                    validation_columns['Found_Details_Count'].append(0)
                    validation_columns['WO_Numbers_Found'].append('')
            
            # Agregar todas las columnas de validación al DataFrame
            for col_name, col_data in validation_columns.items():
                matrix_df[col_name] = col_data
        
        return matrix_df
    
    def export_to_excel(self, matrix_df, results, filename=None):
        """
        Exporta los resultados a un archivo Excel con múltiples hojas
        
        Args:
            matrix_df: DataFrame con la matriz principal
            results: Lista de resultados de validación
            filename: Nombre del archivo (opcional)
            
        Returns:
            str: Ruta del archivo creado
        """
        if filename is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f'Demand_Validation_Results_{timestamp}.xlsx'
        
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                
                # Hoja 1: Matriz principal (Expedite + Validación)
                matrix_df.to_excel(writer, sheet_name='Validation_Matrix', index=False)
                
                # Hoja 2: Resumen por Status
                if results:
                    status_summary = self.create_status_summary(results)
                    status_summary.to_excel(writer, sheet_name='Status_Summary', index=False)
                
                # Hoja 3: Resumen por Tipo de Demanda
                if results:
                    demand_type_summary = self.create_demand_type_summary(results)
                    demand_type_summary.to_excel(writer, sheet_name='Demand_Type_Summary', index=False)
                
                # Hoja 4: Detalles de No Encontrados
                not_found_df = matrix_df[matrix_df['Validation_Status'] == 'NOT_FOUND'].copy()
                if not not_found_df.empty:
                    not_found_df.to_excel(writer, sheet_name='Not_Found_Details', index=False)
                
                # Hoja 5: Detalles de Encontrados
                found_df = matrix_df[matrix_df['Validation_Status'] == 'FOUND'].copy()
                if not found_df.empty:
                    found_df.to_excel(writer, sheet_name='Found_Details', index=False)
                
                # Hoja 6: Estadísticas generales
                if results:
                    stats_df = self.create_general_statistics(results, matrix_df)
                    stats_df.to_excel(writer, sheet_name='General_Statistics', index=False)
            
            logger.info(f"Archivo Excel creado exitosamente: {filename}")
            return filename
            
        except Exception as e:
            logger.error(f"Error al crear archivo Excel: {e}")
            return None
    
    def create_status_summary(self, results):
        """Crea resumen por status de validación"""
        status_counts = {}
        for result in results:
            status = result.get('status', 'UNKNOWN')
            status_counts[status] = status_counts.get(status, 0) + 1
        
        summary_data = []
        total = len(results)
        for status, count in status_counts.items():
            percentage = round((count / total) * 100, 2) if total > 0 else 0
            summary_data.append({
                'Status': status,
                'Count': count,
                'Percentage': f"{percentage}%"
            })
        
        return pd.DataFrame(summary_data)
    
    def create_demand_type_summary(self, results):
        """Crea resumen por tipo de demanda"""
        demand_type_counts = {}
        demand_type_found = {}
        
        for result in results:
            demand_type = result.get('demand_type', 'UNKNOWN')
            status = result.get('status', 'UNKNOWN')
            
            demand_type_counts[demand_type] = demand_type_counts.get(demand_type, 0) + 1
            
            if status == 'FOUND':
                demand_type_found[demand_type] = demand_type_found.get(demand_type, 0) + 1
        
        summary_data = []
        for demand_type, total_count in demand_type_counts.items():
            found_count = demand_type_found.get(demand_type, 0)
            success_rate = round((found_count / total_count) * 100, 2) if total_count > 0 else 0
            
            summary_data.append({
                'Demand_Type': demand_type,
                'Total_Count': total_count,
                'Found_Count': found_count,
                'Not_Found_Count': total_count - found_count,
                'Success_Rate': f"{success_rate}%"
            })
        
        return pd.DataFrame(summary_data)
    
    def create_general_statistics(self, results, matrix_df):
        """Crea estadísticas generales"""
        total_demands = len(results)
        found_demands = len([r for r in results if r.get('status') == 'FOUND'])
        not_found_demands = len([r for r in results if r.get('status') == 'NOT_FOUND'])
        
        # Estadísticas por AC
        ac_stats = matrix_df.groupby('AC').agg({
            'Validation_Status': ['count', lambda x: (x == 'FOUND').sum()],
            'ReqQty': 'sum'
        }).round(2)
        
        ac_stats.columns = ['Total_Demands', 'Found_Demands', 'Total_ReqQty']
        ac_stats['Success_Rate'] = round((ac_stats['Found_Demands'] / ac_stats['Total_Demands']) * 100, 2)
        ac_stats = ac_stats.reset_index()
        
        return ac_stats
    
    def generate_validation_report(self, results):
        """
        Genera un reporte de validación
        
        Args:
            results (list): Lista de resultados de validación
            
        Returns:
            dict: Reporte resumido
        """
        if not results:
            return {'message': 'No hay resultados para reportar'}
        
        # Contar por status
        status_counts = {}
        demand_type_counts = {}
        
        for result in results:
            status = result.get('status', 'UNKNOWN')
            demand_type = result.get('demand_type', 'UNKNOWN')
            
            status_counts[status] = status_counts.get(status, 0) + 1
            demand_type_counts[demand_type] = demand_type_counts.get(demand_type, 0) + 1
        
        # Calcular estadísticas
        total_demands = len(results)
        found_demands = status_counts.get('FOUND', 0)
        not_found_demands = status_counts.get('NOT_FOUND', 0)
        
        report = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'total_demands_validated': total_demands,
            'summary': {
                'found': found_demands,
                'not_found': not_found_demands,
                'success_rate': round((found_demands / total_demands) * 100, 2) if total_demands > 0 else 0
            },
            'status_breakdown': status_counts,
            'demand_type_breakdown': demand_type_counts
        }
        
        return report


def main():
    """Función principal para ejecutar el validador eficiente"""
    # Crear instancia del validador
    validator = DemandValidator()
    
    # Ejecutar validación eficiente
    print("Iniciando validación eficiente de demandas...")
    results_df, excel_filename = validator.validate_all_demands_efficient(export_excel=True)
    
    if not results_df.empty:
        print("\n" + "="*50)
        print("REPORTE DE VALIDACIÓN EFICIENTE")
        print("="*50)
        print(f"Total de demandas validadas: {len(results_df)}")
        
        # Resumen por status
        status_summary = results_df['Validation_Status'].value_counts()
        print("\nResumen por status:")
        for status, count in status_summary.items():
            percentage = (count / len(results_df)) * 100
            print(f"  {status}: {count} ({percentage:.1f}%)")
        
        # Resumen por tipo de demanda
        demand_summary = results_df.groupby(['DemandType', 'Validation_Status']).size().unstack(fill_value=0)
        print("\nResumen por tipo de demanda:")
        for demand_type in demand_summary.index:
            found = demand_summary.loc[demand_type, 'FOUND'] if 'FOUND' in demand_summary.columns else 0
            not_found = demand_summary.loc[demand_type, 'NOT_FOUND'] if 'NOT_FOUND' in demand_summary.columns else 0
            total = found + not_found
            success_rate = (found / total * 100) if total > 0 else 0
            print(f"  {demand_type}: {found}/{total} encontradas ({success_rate:.1f}%)")
        
        # Mostrar información del archivo Excel
        if excel_filename:
            full_path = os.path.abspath(excel_filename)
            print(f"\n📊 ARCHIVO EXCEL GENERADO:")
            print(f"   Archivo: {excel_filename}")
            print(f"   Ruta completa: {full_path}")
            print(f"   Hojas incluidas:")
            print(f"     • Validation_Results: Resultados principales")
            print(f"     • Status_Summary: Resumen por status y tipo")
            print(f"     • AC_Summary: Resumen por AC")
            print(f"     • Not_Found: Solo demandas no encontradas")
            print(f"     • Found: Solo demandas encontradas")
        
        # Tiempo de procesamiento estimado
        total_records = len(results_df)
        print(f"\n⚡ EFICIENCIA:")
        print(f"   Procesadas {total_records:,} demandas usando merges eficientes")
        print(f"   Mucho más rápido que validación row-by-row")
    
    else:
        print("No se pudieron obtener resultados de validación")


if __name__ == "__main__":
    main()