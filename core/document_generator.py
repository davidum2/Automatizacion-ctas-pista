import os
import logging
from pathlib import Path

# Importar las funciones específicas de cada módulo
from generators.creacionDocumentos import creacionDocumentos
from utils.web_utils import descargar_verificacion

class DocumentGenerator:
    """
    Clase para generar documentos Word a partir de datos XML procesados.
    Esta versión solo genera archivos DOCX, sin conversión a PDF.
    """
    
    def __init__(self):
        """Inicializa el generador de documentos."""
        self.logger = logging.getLogger(__name__)
        
        # Configurar el logging básico si no está configurado
        if not logging.getLogger().handlers:
            logging.basicConfig(level=logging.INFO, 
                                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    def generate_all_documents(self, data, output_dir):
        """
        Genera todos los documentos DOCX para una factura.
        
        Args:
            data (dict): Datos extraídos del XML
            output_dir (str): Directorio donde se guardarán los documentos
            
        Returns:
            dict: Diccionario con las rutas a los documentos generados
        """


        try:
            descargar_verificacion(data, output_dir)
        
        except Exception as e:
            self.logger.error(f"Error general en la generación de documentos: {str(e)}")
        




        try:
            # Verificar que el directorio de salida existe
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Crear diccionario para guardar las rutas de los documentos generados
            generated_files = {}
            
            # Lista de nombres de plantillas
            template_files = [
                'legalizacion_factura.docx',
                'legalizacion_verificacion.docx',
                'legalizacion_xml.docx',
                'xml.docx'
            ]
            
            # Directorio base de plantillas
            templates_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "plantillas")
            
            # Procesar cada plantilla
            for template_file in template_files:
                template_path = os.path.join(templates_dir, template_file)
                template_name = template_file.replace('.docx', '').replace('_', ' ')
                
                self.logger.info(f"Generando {template_name}...")
                try:
                    # Verificar que la plantilla existe
                    if not os.path.exists(template_path):
                        self.logger.error(f"No se encontró la plantilla: {template_path}")
                        continue
                        
                    # Generar el documento usando la misma función para todas las plantillas
                    generated_file = creacionDocumentos(template_path, output_dir, data, template_name)
                    generated_files[template_file.replace('.docx', '')] = generated_file
                    self.logger.info(f"✓ {template_name.capitalize()} generado correctamente")
                    
                except Exception as e:
                    self.logger.error(f"Error al generar {template_name}: {str(e)}")
            
            return generated_files

        except Exception as e:
            self.logger.error(f"Error general en la generación de documentos: {str(e)}")
            raise