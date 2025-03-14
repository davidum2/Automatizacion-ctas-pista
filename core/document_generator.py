"""
Clase para generar documentos Word y PDF a partir de datos XML procesados.
"""
import os
import logging
from pathlib import Path

# Importar las funciones específicas de cada módulo
from generators.creacionDocumentos import creacionDocumentos
from utils.web_utils import descargar_verificacion
from factura_pdf_processor import FacturaPDFProcessor 

class DocumentGenerator:
    """
    Clase para generar documentos Word y PDF a partir de datos XML procesados.
    Genera archivos DOCX y los convierte a PDF, además de combinar PDFs según el formato requerido.
    """
    
    def __init__(self, ui=None):
        """
        Inicializa el generador de documentos.
        
        Args:
            ui: Referencia opcional a la interfaz de usuario para reportes
        """
        self.ui = ui
        self.logger = logging.getLogger(__name__)
        self.pdf_processor = FacturaPDFProcessor(ui)
        
        # Configurar el logging básico si no está configurado
        if not logging.getLogger().handlers:
            logging.basicConfig(level=logging.INFO, 
                                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    def update_status(self, message, level="info"):
        """
        Actualiza el estado en la UI si está disponible.
        
        Args:
            message (str): Mensaje a mostrar
            level (str): Nivel del mensaje (info, warning, error, success)
        """
        if self.ui and hasattr(self.ui, 'update_status'):
            self.ui.update_status(message, level)
        
        # También registrar en el log
        if level == "error":
            self.logger.error(message)
        elif level == "warning":
            self.logger.warning(message)
        else:
            self.logger.info(message)
            
    def generate_docx_documents(self, data, output_dir):
        """
        Genera documentos DOCX para una factura.
        
        Args:
            data (dict): Datos extraídos del XML
            output_dir (str): Directorio donde se guardarán los documentos
            
        Returns:
            dict: Diccionario con las rutas a los documentos generados
        """
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
                'legalizacion_xmls.docx',
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
            self.logger.error(f"Error general en la generación de documentos DOCX: {str(e)}")
            raise
            
    def generate_all_documents(self, data, output_dir):
        """
        Genera todos los documentos para una factura (DOCX y PDF).
        
        Args:
            data (dict): Datos extraídos del XML
            output_dir (str): Directorio donde se guardarán los documentos
            
        Returns:
            dict: Diccionario con las rutas a los documentos generados
        """
        try:
            # Paso 1: Descargar verificación del SAT si está configurado
            self.update_status("Intentando descargar verificación del SAT...")
            try:
                descargar_verificacion(data, output_dir)
            except Exception as e:
                self.logger.warning(f"No se pudo descargar verificación del SAT: {str(e)}")
            
            # Paso 2: Generar documentos DOCX
            self.update_status("Generando documentos Word...")
            docx_files = self.generate_docx_documents(data, output_dir)
            
            if not docx_files:
                self.update_status("No se generaron documentos Word", "error")
                return {}
            
            # Paso 3: Procesar PDFs
            self.update_status("Procesando documentos PDF...")
            
            # Obtener la ruta del XML original
            xml_dir = os.path.dirname(data.get('xml_path', ''))
            if not xml_dir:
                # Si no tenemos la ruta del XML, asumimos que está en el directorio de salida
                xml_dir = output_dir
                
            # Crear directorio para PDFs si no existe
            pdf_dir = os.path.join(output_dir, "pdfs")
            if not os.path.exists(pdf_dir):
                os.makedirs(pdf_dir)
                
            # Procesar PDFs para generar documento combinado
            pdf_results = self.pdf_processor.process_factura_pdfs(
                os.path.join(xml_dir, "factura.xml"),  # Asumimos este nombre si no tenemos la ruta real
                pdf_dir,
                docx_files
            )
            
            # Combinar resultados de documentos DOCX y PDF
            results = {
                'docx_files': docx_files,
                'pdf_processed': False,
                'pdf_files': {}
            }
            
            if pdf_results:
                results['pdf_processed'] = True
                results['pdf_files'] = pdf_results
                self.update_status("Procesamiento de PDFs completado con éxito", "success")
            else:
                self.update_status("No se completó el procesamiento de PDFs", "warning")
            
            return results

        except Exception as e:
            self.logger.error(f"Error general en la generación de documentos: {str(e)}")
            import traceback
            traceback.print_exc()
            return {'error': str(e)}