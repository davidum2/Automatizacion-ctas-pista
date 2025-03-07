import pandas as pd
import os
import logging

class ExcelReader:
    """
    Clase para leer archivos Excel con formato específico.
    El archivo Excel debe tener una hoja llamada 'base datos' con columnas:
    'PARTIDA', 'CONCEPTO', 'MONTO', 'NUMERO' (en mayúsculas)
    """
    
    def __init__(self):
        """Inicializa el lector de Excel con configuración predeterminada."""
        # Configurar logging
        self.logger = logging.getLogger(__name__)
        
        # Nombre de la hoja que contiene los datos
        self.sheet_name = 'base datos'
        
        # Nombres de columnas esperados
        self.expected_columns = ['PARTIDA', 'CONCEPTO', 'MONTO', 'NUMERO']
    
    def read_partidas(self, excel_path):
        """
        Lee el archivo Excel con formato específico para extraer información de partidas.
        
        Args:
            excel_path (str): Ruta al archivo Excel
            
        Returns:
            list: Lista de diccionarios con detalles de partidas
        """
        try:
            # Verificar que el archivo existe
            if not os.path.exists(excel_path):
                raise FileNotFoundError(f"No se encontró el archivo: {excel_path}")
            
            # Verificar que el archivo Excel tiene la hoja esperada
            try:
                xls = pd.ExcelFile(excel_path)
                if self.sheet_name not in xls.sheet_names:
                    self.logger.error(f"El archivo Excel no contiene la hoja '{self.sheet_name}'")
                    self.logger.info(f"Hojas disponibles: {xls.sheet_names}")
                    raise ValueError(f"No se encontró la hoja '{self.sheet_name}' en el archivo Excel")
            except Exception as e:
                self.logger.error(f"Error al verificar las hojas del Excel: {str(e)}")
                raise
            
            # Leer el archivo Excel directamente con los nombres de columnas esperados
            self.logger.info(f"Leyendo hoja '{self.sheet_name}' del archivo Excel...")
            df = pd.read_excel(excel_path, sheet_name=self.sheet_name)
            
            # Verificar que todas las columnas esperadas están presentes
            missing_columns = [col for col in self.expected_columns if col not in df.columns]
            if missing_columns:
                self.logger.error(f"Faltan columnas en el Excel: {missing_columns}")
                self.logger.info(f"Columnas encontradas: {df.columns.tolist()}")
                raise ValueError(f"El archivo Excel no contiene las columnas requeridas: {missing_columns}")
            
            # Limpiar datos: eliminar filas vacías y convertir NaN a None
            df = df.dropna(subset=['PARTIDA', 'MONTO'], how='any')
            df = df.fillna({'CONCEPTO': '', 'NUMERO': ''})
            
            # Crear lista de partidas
            partidas = []
            for _, row in df.iterrows():
                try:
                    # Convertir partida a string y normalizar (eliminar decimales si es necesario)
                    partida_value = row['PARTIDA']
                    if isinstance(partida_value, (int, float)):
                        partida_num = str(int(partida_value))
                    else:
                        partida_num = str(partida_value).strip()
                    
                    # Obtener concepto/descripción
                    descripcion = str(row['CONCEPTO']).strip()
                    if not descripcion:
                        descripcion = f"Partida {partida_num}"
                    
                    # Obtener y validar monto
                    monto = row['MONTO']
                    if not isinstance(monto, (int, float)) or pd.isna(monto):
                        self.logger.warning(f"Valor de monto no válido para partida {partida_num}: {monto}")
                        continue
                    
                    # Obtener número (opcional)
                    numero = row['NUMERO']
                    if pd.isna(numero):
                        numero = ""
                    elif isinstance(numero, (int, float)):
                        numero = str(int(numero)) if numero == int(numero) else str(numero)
                    else:
                        numero = str(numero).strip()
                    
                    # Crear objeto partida
                    partida = {
                        'numero': partida_num,
                        'descripcion': descripcion,
                        'monto': float(monto),
                        'numero_adicional': numero
                    }
                    partidas.append(partida)
                    
                except Exception as e:
                    self.logger.error(f"Error al procesar fila {_}: {str(e)}")
                    # Continuar con la siguiente fila
            
            self.logger.info(f"Se encontraron {len(partidas)} partidas en el archivo")
            
            # Si no se encontraron partidas, mostrar advertencia
            if not partidas:
                self.logger.warning("No se encontraron partidas válidas en el archivo")
                raise ValueError("No se encontraron partidas válidas en el archivo Excel")
            
            return partidas
            
        except pd.errors.EmptyDataError:
            self.logger.error("El archivo Excel está vacío")
            raise Exception("El archivo Excel está vacío")
        except pd.errors.ParserError:
            self.logger.error("Error al analizar el archivo Excel. Formato no válido.")
            raise Exception("Error al analizar el archivo Excel. Formato no válido.")
        except Exception as e:
            self.logger.error(f"Error al leer el archivo Excel: {str(e)}")
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