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
    
    def validate_so_demand(self, expedite_row):
        """
        Valida demanda tipo SO (placeholder - necesitas especificar la tabla y lógica)
        
        Args:
            expedite_row: Fila de expedite a validar
            
        Returns:
            dict: Resultado de la validación
        """
        # TODO: Implementar validación para SO cuando especifiques la tabla/lógica
        return {
            'status': 'PENDING_IMPLEMENTATION',
            'source_table': 'SO',
            'message': 'Validación SO pendiente de implementación'
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
    
    def validate_single_demand(self, expedite_row, fcst_df, wo_df):
        """
        Valida una sola demanda según su tipo
        
        Args:
            expedite_row: Fila de expedite a validar
            fcst_df: DataFrame con datos de fcst
            wo_df: DataFrame con datos de WorkOrder
            
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
                validation_result.update(self.validate_so_demand(expedite_row))
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
    
    def validate_all_demands(self, export_excel=True):
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
            
            if expedite_df.empty:
                logger.warning("No se encontraron datos en expedite")
                return [], None
            
            # Validar cada demanda
            results = []
            logger.info(f"Iniciando validación de {len(expedite_df)} demandas...")
            
            for index, row in expedite_df.iterrows():
                result = self.validate_single_demand(row, fcst_df, wo_df)
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
                'Qty_Difference': [],
                'Qty_Status': [],
                'Validation_Message': [],
                'Found_Details_Count': [],
                'WO_Numbers_Found': [],
                'FCST_Number_Found': []
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
                validation_columns['Qty_Difference'].append(result.get('qty_difference', ''))
                validation_columns['Qty_Status'].append(result.get('qty_status', ''))
                validation_columns['FCST_Number_Found'].append(result.get('fcst_number', ''))
                
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
    """Función principal para ejecutar el validador"""
    # Crear instancia del validador
    validator = DemandValidator()
    
    # Ejecutar validación
    print("Iniciando validación de demandas...")
    results, excel_filename = validator.validate_all_demands(export_excel=True)
    
    if results:
        # Generar reporte
        report = validator.generate_validation_report(results)
        
        print("\n" + "="*50)
        print("REPORTE DE VALIDACIÓN DE DEMANDAS")
        print("="*50)
        print(f"Timestamp: {report['timestamp']}")
        print(f"Total de demandas validadas: {report['total_demands_validated']}")
        print(f"Encontradas: {report['summary']['found']}")
        print(f"No encontradas: {report['summary']['not_found']}")
        print(f"Tasa de éxito: {report['summary']['success_rate']}%")
        
        print("\nDesglose por tipo de demanda:")
        for demand_type, count in report['demand_type_breakdown'].items():
            print(f"  {demand_type}: {count}")
        
        print("\nDesglose por status:")
        for status, count in report['status_breakdown'].items():
            print(f"  {status}: {count}")
        
        # Mostrar información del archivo Excel
        if excel_filename:
            full_path = os.path.abspath(excel_filename)
            print(f"\n📊 ARCHIVO EXCEL GENERADO:")
            print(f"   Archivo: {excel_filename}")
            print(f"   Ruta completa: {full_path}")
            print(f"   Hojas incluidas:")
            print(f"     • Validation_Matrix: Matriz principal (expedite + validación)")
            print(f"     • Status_Summary: Resumen por status")
            print(f"     • Demand_Type_Summary: Resumen por tipo de demanda")
            print(f"     • Not_Found_Details: Detalles de no encontrados")
            print(f"     • Found_Details: Detalles de encontrados")
            print(f"     • General_Statistics: Estadísticas por AC")
        
        # Opcional: Mostrar algunos ejemplos de demandas no encontradas
        not_found_examples = [r for r in results if r.get('status') == 'NOT_FOUND'][:5]
        if not_found_examples:
            print(f"\nEjemplos de demandas no encontradas (primeras 5):")
            for example in not_found_examples:
                print(f"  Key: {example['expedite_key']}, ItemNo: {example['item_no']}, Type: {example['demand_type']}")
    
    else:
        print("No se pudieron obtener resultados de validación")


if __name__ == "__main__":
    main()