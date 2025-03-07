import pandas as pd
import numpy as np
import os
import logging

class ExcelReader:
    """
    Clase para leer archivos Excel con formato específico donde los encabezados
    están en la fila 6 (celdas B6-D6) y los datos comienzan en la fila 7.
    """
    
    def __init__(self):
        """Inicializa el lector de Excel con configuración predeterminada."""
        # Configurar logging
        self.logger = logging.getLogger(__name__)
        
        # Mapeo de nombres de columna alternativos que pueden estar en el Excel
        self.column_mappings = {
            'partida': ['Partida', 'PARTIDA', 'partida', 'Clave', 'CLAVE', 'Número', 'No.', 'Num'],
            'descripcion': ['Descripcion', 'DESCRIPCION', 'Descripción', 'DESCRIPCIÓN', 'Concepto', 'CONCEPTO', 'Detalle'],
            'monto': ['Monto', 'MONTO', 'Importe', 'IMPORTE', 'Total', 'TOTAL', 'Cantidad', 'Presupuesto']
        }
    
    def read_partidas(self, excel_path, header_row=5, data_start_row=6, sheet_name=0):
        """
        Lee el archivo Excel con formato específico para extraer información de partidas.
        
        Args:
            excel_path (str): Ruta al archivo Excel
            header_row (int): Índice de la fila que contiene los encabezados (0-based, por defecto 5 para fila 6)
            data_start_row (int): Índice de la fila donde comienzan los datos (0-based, por defecto 6 para fila 7)
            sheet_name (str o int, opcional): Nombre u índice de la hoja a leer. Por defecto 0 (primera hoja).
            
        Returns:
            list: Lista de diccionarios con detalles de partidas
        """
        try:
            # Verificar que el archivo existe
            if not os.path.exists(excel_path):
                raise FileNotFoundError(f"No se encontró el archivo: {excel_path}")
            
            # Leer el archivo Excel sin encabezados primero para examinar su estructura
            df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
            
            # Extraer los encabezados de la fila especificada
            if header_row >= len(df_raw):
                self.logger.warning(f"La fila de encabezados {header_row+1} está fuera del rango. El Excel solo tiene {len(df_raw)} filas.")
                # Intentar detectar encabezados automáticamente
                for i in range(min(10, len(df_raw))):
                    self.logger.info(f"Fila {i+1}: {df_raw.iloc[i].tolist()}")
                raise ValueError(f"La fila de encabezados {header_row+1} no existe en el archivo")
            
            # Mostrar información de depuración
            self.logger.info(f"Posibles encabezados en fila {header_row+1}: {df_raw.iloc[header_row].tolist()}")
            
            # Leer nuevamente el archivo, ahora especificando la fila de encabezados
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row, skiprows=lambda x: x < header_row or (x > header_row and x < data_start_row))
            
            # Mostrar columnas encontradas
            self.logger.info(f"Columnas encontradas: {df.columns.tolist()}")
            
            # Si el Excel tiene valores NaN en los encabezados, pandas los nombra como Unnamed: X
            # Vamos a buscar las columnas relevantes por posición también
            column_indices = {'partida': None, 'descripcion': None, 'monto': None}
            column_found = {'partida': False, 'descripcion': False, 'monto': False}
            
            # Buscar por nombre primero
            for req_col, possible_names in self.column_mappings.items():
                for col in df.columns:
                    if any(possible in str(col).lower() for possible in [p.lower() for p in possible_names]):
                        column_indices[req_col] = col
                        column_found[req_col] = True
                        self.logger.info(f"Columna '{req_col}' encontrada como '{col}'")
                        break
            
            # Si no se encontraron todas las columnas por nombre, intentar inferir por posición
            # Asumiendo que las columnas están en el orden: partida, descripción, monto
            if not all(column_found.values()):
                columns_order = ['partida', 'descripcion', 'monto']
                unnamed_columns = [col for col in df.columns if 'Unnamed:' in str(col)]
                
                if len(unnamed_columns) + sum(column_found.values()) >= 3:
                    index = 0
                    for req_col in columns_order:
                        if not column_found[req_col]:
                            if index < len(unnamed_columns):
                                column_indices[req_col] = unnamed_columns[index]
                                self.logger.info(f"Columna '{req_col}' inferida como '{unnamed_columns[index]}'")
                                index += 1
                            else:
                                # Si no hay suficientes columnas sin nombre, usar columnas basadas en índice
                                numeric_cols = [i for i, col in enumerate(df.columns) if not column_found[col]]
                                if numeric_cols:
                                    col_idx = numeric_cols[0]
                                    column_indices[req_col] = df.columns[col_idx]
                                    self.logger.info(f"Columna '{req_col}' asignada por posición a '{df.columns[col_idx]}'")
            
            # Verificar que tenemos todas las columnas necesarias
            missing_cols = [col for col, found in column_found.items() if not found and column_indices[col] is None]
            if missing_cols:
                # Si no podemos encontrar automáticamente, intentar con índices fijos (B, C, D)
                if len(df.columns) >= 3:
                    partida_col = df.columns[1]  # Columna B (índice 1)
                    desc_col = df.columns[2]     # Columna C (índice 2)
                    monto_col = df.columns[3]    # Columna D (índice 3)
                    
                    self.logger.warning(f"Usando asignación fija de columnas: {partida_col}, {desc_col}, {monto_col}")
                    
                    column_indices['partida'] = partida_col
                    column_indices['descripcion'] = desc_col
                    column_indices['monto'] = monto_col
                else:
                    self.logger.error(f"No se pudieron encontrar las columnas: {missing_cols}")
                    self.logger.error(f"Columnas disponibles: {df.columns.tolist()}")
                    raise ValueError(f"No se pudieron encontrar todas las columnas necesarias: {missing_cols}")
            
            # Limpiar y preparar datos
            df = df.replace({np.nan: None})
            df = df.dropna(how='all')
            
            # Crear lista de partidas
            partidas = []
            
            for _, row in df.iterrows():
                partida_value = row[column_indices['partida']]
                descripcion_value = row[column_indices['descripcion']]
                monto_value = row[column_indices['monto']]
                
                # Omitir filas sin valores significativos
                if pd.isna(partida_value) or pd.isna(monto_value):
                    continue
                
                # Convertir partida a string y normalizar
                try:
                    partida_num = str(partida_value)
                    if partida_num.replace('.', '', 1).isdigit():
                        partida_num = str(int(float(partida_num)))
                except:
                    # Si hay error al convertir, usar el valor tal cual
                    partida_num = str(partida_value)
                
                # Si descripción es NaN, usar un valor por defecto
                if pd.isna(descripcion_value):
                    descripcion_value = f"Partida {partida_num}"
                
                # Convertir monto a float con manejo de errores
                try:
                    monto_float = float(monto_value)
                except (ValueError, TypeError):
                    # Si no se puede convertir, intentar limpiar el string y volver a intentar
                    try:
                        monto_str = str(monto_value).replace(',', '').replace('$', '').strip()
                        monto_float = float(monto_str)
                    except:
                        self.logger.warning(f"No se pudo convertir el monto '{monto_value}' para la partida {partida_num}")
                        continue
                
                partida = {
                    'numero': partida_num,
                    'descripcion': str(descripcion_value),
                    'monto': monto_float
                }
                partidas.append(partida)
            
            self.logger.info(f"Se encontraron {len(partidas)} partidas en el archivo")
            
            # Si no se encontraron partidas, podría ser un problema con las filas de encabezado
            if not partidas:
                self.logger.warning("No se encontraron partidas. Mostrando primeras 10 filas del Excel para depuración:")
                for i in range(min(10, len(df_raw))):
                    self.logger.info(f"Fila {i+1}: {df_raw.iloc[i].tolist()}")
                raise ValueError("No se encontraron partidas válidas en el archivo")
            
            return partidas
            
        except pd.errors.EmptyDataError:
            raise Exception("El archivo Excel está vacío")
        except pd.errors.ParserError:
            raise Exception("Error al analizar el archivo Excel. Formato no válido.")
        except Exception as e:
            raise Exception(f"Error al leer el archivo Excel: {str(e)}")

    def get_available_sheets(self, excel_path):
        """
        Obtiene los nombres de todas las hojas disponibles en el archivo Excel.
        
        Args:
            excel_path (str): Ruta al archivo Excel
            
        Returns:
            list: Lista con los nombres de las hojas
        """
        try:
            xls = pd.ExcelFile(excel_path)
            return xls.sheet_names
        except Exception as e:
            self.logger.error(f"Error al obtener hojas del Excel: {str(e)}")
            return []
    
    def analyze_excel_structure(self, excel_path, sheet_name=0, max_rows=15):
        """
        Analiza la estructura del Excel para ayudar a identificar encabezados y datos.
        Útil para depuración.
        
        Args:
            excel_path (str): Ruta al archivo Excel
            sheet_name (str o int): Hoja a analizar
            max_rows (int): Número máximo de filas a mostrar
            
        Returns:
            dict: Información sobre la estructura del Excel
        """
        try:
            # Leer sin encabezados
            df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
            
            # Obtener información básica
            info = {
                "total_filas": len(df_raw),
                "total_columnas": len(df_raw.columns),
                "muestra_filas": []
            }
            
            # Mostrar contenido de las primeras filas
            for i in range(min(max_rows, len(df_raw))):
                row_content = {}
                for j, col in enumerate(df_raw.columns):
                    cell_value = df_raw.iloc[i, j]
                    # Convertir NaN a None para mejor visualización
                    if pd.isna(cell_value):
                        cell_value = None
                    row_content[f"Col_{j+1}"] = cell_value
                info["muestra_filas"].append({"fila": i+1, "contenido": row_content})
            
            return info
            
        except Exception as e:
            return {"error": str(e)}