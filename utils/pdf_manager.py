"""
Clase para gestionar operaciones con PDFs, incluyendo conversión, combinación y manipulación
"""
import os
import logging
from PyPDF2 import PdfReader, PdfWriter
import pikepdf
import tempfile
import shutil
from docx2pdf import convert as docx2pdf_convert

# Configurar logging
logger = logging.getLogger(__name__)

class PDFManager:
    """
    Clase para gestionar operaciones con archivos PDF.
    Proporciona funcionalidades para convertir documentos DOCX a PDF,
    combinar varios PDFs y manipular páginas de PDF.
    """
    
    def __init__(self):
        """Inicializa el gestor de PDFs."""
        self.temp_dir = None
        self._create_temp_dir()
    
    def _create_temp_dir(self):
        """Crea un directorio temporal para operaciones con PDF."""
        try:
            self.temp_dir = tempfile.mkdtemp(prefix="pdf_manager_")
            logger.debug(f"Directorio temporal creado: {self.temp_dir}")
        except Exception as e:
            logger.error(f"Error al crear directorio temporal: {str(e)}")
            raise
    
    def convert_docx_to_pdf(self, docx_path, output_dir=None):
        """
        Convierte un archivo DOCX a PDF.
        
        Args:
            docx_path (str): Ruta al archivo DOCX
            output_dir (str, optional): Directorio donde guardar el PDF.
                                        Si es None, se usa el mismo directorio que el DOCX.
                                        
        Returns:
            str: Ruta al archivo PDF generado
        """
        if output_dir is None:
            output_dir = os.path.dirname(docx_path)
        
        try:
            # Obtener nombre del archivo sin extensión
            file_name = os.path.basename(docx_path)
            name_without_ext = os.path.splitext(file_name)[0]
            
            # Generar ruta de salida
            pdf_path = os.path.join(output_dir, f"{name_without_ext}.pdf")
            
            # Convertir DOCX a PDF
            logger.info(f"Convirtiendo {docx_path} a PDF...")
            docx2pdf_convert(docx_path, pdf_path)
            
            # Verificar que el archivo PDF se creó correctamente
            if os.path.exists(pdf_path):
                logger.info(f"PDF generado exitosamente: {pdf_path}")
                return pdf_path
            else:
                raise FileNotFoundError(f"No se generó el archivo PDF: {pdf_path}")
        
        except Exception as e:
            logger.error(f"Error al convertir DOCX a PDF: {str(e)}")
            raise
    
    def convert_multiple_docx(self, docx_paths, output_dir=None):
        """
        Convierte múltiples archivos DOCX a PDF.
        
        Args:
            docx_paths (list): Lista de rutas a archivos DOCX
            output_dir (str, optional): Directorio donde guardar los PDFs.
                                        Si es None, se usa el mismo directorio que cada DOCX.
                                        
        Returns:
            dict: Diccionario con rutas a los archivos PDF generados
        """
        pdf_paths = {}
        
        for docx_path in docx_paths:
            try:
                # Determinar directorio de salida
                output_path = self.convert_docx_to_pdf(docx_path, output_dir)
                
                # Guardar ruta en el diccionario
                base_name = os.path.basename(docx_path)
                name_without_ext = os.path.splitext(base_name)[0]
                pdf_paths[name_without_ext] = output_path
                
            except Exception as e:
                logger.error(f"Error al convertir {docx_path}: {str(e)}")
                # Continuar con el siguiente archivo
        
        return pdf_paths
    
    def count_pdf_pages(self, pdf_path):
        """
        Cuenta el número de páginas en un archivo PDF.
        
        Args:
            pdf_path (str): Ruta al archivo PDF
            
        Returns:
            int: Número de páginas
        """
        try:
            with open(pdf_path, 'rb') as file:
                pdf = PdfReader(file)
                num_pages = len(pdf.pages)
                logger.debug(f"PDF {pdf_path} tiene {num_pages} páginas")
                return num_pages
        except Exception as e:
            logger.error(f"Error al contar páginas del PDF {pdf_path}: {str(e)}")
            raise
    
    def combine_pdfs(self, output_path, pdf_files):
        """
        Combina múltiples archivos PDF en uno solo.
        
        Args:
            output_path (str): Ruta donde guardar el PDF combinado
            pdf_files (list): Lista de rutas a archivos PDF para combinar
            
        Returns:
            str: Ruta al archivo PDF combinado
        """
        try:
            pdf_writer = PdfWriter()
            
            # Añadir cada PDF
            for pdf_file in pdf_files:
                if not os.path.exists(pdf_file):
                    logger.warning(f"Archivo PDF no encontrado: {pdf_file}")
                    continue
                
                with open(pdf_file, 'rb') as file:
                    pdf_reader = PdfReader(file)
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)
            
            # Guardar el PDF combinado
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            
            logger.info(f"PDFs combinados exitosamente en: {output_path}")
            return output_path
        
        except Exception as e:
            logger.error(f"Error al combinar PDFs: {str(e)}")
            raise
    
    def create_alternating_pdf(self, output_path, main_pdf, interleaved_pdf):
        """
        Crea un PDF donde cada página del documento principal es seguida
        por una página del documento intercalado.
        
        Args:
            output_path (str): Ruta donde guardar el PDF resultante
            main_pdf (str): Ruta al PDF principal
            interleaved_pdf (str): Ruta al PDF que se intercalará
            
        Returns:
            str: Ruta al PDF resultante
        """
        try:
            # Abrir PDFs
            with open(main_pdf, 'rb') as main_file, open(interleaved_pdf, 'rb') as interleaved_file:
                main_reader = PdfReader(main_file)
                interleaved_reader = PdfReader(interleaved_file)
                
                # Comprobar que hay al menos una página en cada documento
                if len(main_reader.pages) == 0 or len(interleaved_reader.pages) == 0:
                    raise ValueError("Ambos PDFs deben tener al menos una página")
                
                # Si el documento intercalado tiene solo una página, la usaremos varias veces
                single_interleaved_page = len(interleaved_reader.pages) == 1
                
                # Crear el nuevo PDF
                writer = PdfWriter()
                
                # Añadir páginas alternadas
                for i in range(len(main_reader.pages)):
                    # Añadir página del documento principal
                    writer.add_page(main_reader.pages[i])
                    
                    # Añadir página del documento intercalado
                    if single_interleaved_page:
                        # Siempre usar la primera página si solo hay una
                        writer.add_page(interleaved_reader.pages[0])
                    else:
                        # Intentar usar la página correspondiente o la última disponible
                        interleaved_index = min(i, len(interleaved_reader.pages) - 1)
                        writer.add_page(interleaved_reader.pages[interleaved_index])
                
                # Guardar el resultado
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                
                logger.info(f"PDF alternado creado exitosamente: {output_path}")
                return output_path
                
        except Exception as e:
            logger.error(f"Error al crear PDF alternado: {str(e)}")
            raise
    
    def create_complex_document(self, output_path, document_config):
        """
        Crea un documento PDF complejo siguiendo una configuración específica.
        
        Args:
            output_path (str): Ruta donde guardar el PDF resultante
            document_config (list): Lista de diccionarios con la configuración de cada documento.
                Cada diccionario debe tener:
                - 'path': Ruta al archivo PDF
                - 'all_pages': True para incluir todas las páginas, False para páginas específicas
                - 'pages': Lista de números de página a incluir (solo si all_pages es False)
                - 'interleave_with': Opcional, ruta a un PDF para intercalar después de cada página
                - 'interleave_once': Opcional, True para intercalar solo después del documento completo
            
        Returns:
            str: Ruta al PDF resultante
        """
        try:
            # Crear el nuevo PDF
            writer = PdfWriter()
            
            # Procesar cada documento en la configuración
            for doc_config in document_config:
                # Verificar que el archivo existe
                if not os.path.exists(doc_config['path']):
                    logger.warning(f"Archivo no encontrado: {doc_config['path']}")
                    continue
                
                # Abrir el PDF
                with open(doc_config['path'], 'rb') as file:
                    reader = PdfReader(file)
                    
                    # Determinar qué páginas incluir
                    if doc_config.get('all_pages', True):
                        pages_to_add = range(len(reader.pages))
                    else:
                        # Ajustar índices a base 0 (los números de página comienzan en 1)
                        pages_to_add = [p-1 for p in doc_config.get('pages', []) if 1 <= p <= len(reader.pages)]
                    
                    # Documento para intercalar
                    interleave_pdf = None
                    interleave_reader = None
                    
                    if 'interleave_with' in doc_config and os.path.exists(doc_config['interleave_with']):
                        with open(doc_config['interleave_with'], 'rb') as interleave_file:
                            interleave_reader = PdfReader(interleave_file)
                            
                            # Si no hay páginas en el documento para intercalar, omitirlo
                            if len(interleave_reader.pages) == 0:
                                interleave_reader = None
                    
                    # Añadir páginas con intercalado si es necesario
                    for i, page_index in enumerate(pages_to_add):
                        # Añadir página del documento principal
                        writer.add_page(reader.pages[page_index])
                        
                        # Intercalar después de cada página si está configurado
                        if interleave_reader and not doc_config.get('interleave_once', False):
                            # Si hay una sola página para intercalar, usarla para todas
                            if len(interleave_reader.pages) == 1:
                                writer.add_page(interleave_reader.pages[0])
                            else:
                                # Intentar usar páginas correspondientes o la última disponible
                                interleave_index = min(i, len(interleave_reader.pages) - 1)
                                writer.add_page(interleave_reader.pages[interleave_index])
                    
                    # Intercalar una sola vez después de todo el documento si está configurado
                    if interleave_reader and doc_config.get('interleave_once', False):
                        writer.add_page(interleave_reader.pages[0])
            
            # Guardar el resultado
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
            logger.info(f"Documento complejo creado exitosamente: {output_path}")
            return output_path
                
        except Exception as e:
            logger.error(f"Error al crear documento complejo: {str(e)}")
            raise
    
    def create_factura_legal_document(self, output_path, factura_pdf, legalizacion_factura_pdf, 
                                      verificacion_sat_pdf, legalizacion_verificacion_pdf,
                                      xml_pdf, legalizacion_xml_pdf):
        """
        Crea un documento PDF legal para una factura siguiendo un formato específico.
        
        Args:
            output_path (str): Ruta donde guardar el PDF resultante
            factura_pdf (str): Ruta al PDF de la factura original
            legalizacion_factura_pdf (str): Ruta al PDF de legalización de factura
            verificacion_sat_pdf (str): Ruta al PDF de verificación del SAT
            legalizacion_verificacion_pdf (str): Ruta al PDF de legalización de verificación
            xml_pdf (str): Ruta al PDF del XML
            legalizacion_xml_pdf (str): Ruta al PDF de legalización del XML
            
        Returns:
            str: Ruta al PDF combinado
        """
        try:
            # Configuración del documento
            document_config = [
                # 1. Factura original con legalización después de cada página
                {
                    'path': factura_pdf,
                    'all_pages': True,
                    'interleave_with': legalizacion_factura_pdf,
                    'interleave_once': False
                },
                # 2. Verificación del SAT
                {
                    'path': verificacion_sat_pdf,
                    'all_pages': True,
                    'interleave_with': legalizacion_verificacion_pdf,
                    'interleave_once': True
                },
                # 3. XML con legalización después de cada página
                {
                    'path': xml_pdf,
                    'all_pages': True,
                    'interleave_with': legalizacion_xml_pdf,
                    'interleave_once': False
                }
            ]
            
            # Crear el documento complejo
            return self.create_complex_document(output_path, document_config)
        
        except Exception as e:
            logger.error(f"Error al crear documento legal de factura: {str(e)}")
            raise
    
    def rotate_pdf_pages(self, pdf_path, output_path, rotation_angle=90):
        """
        Rota las páginas de un PDF.
        
        Args:
            pdf_path (str): Ruta al archivo PDF original
            output_path (str): Ruta donde guardar el PDF rotado
            rotation_angle (int): Ángulo de rotación en grados (90, 180, 270)
            
        Returns:
            str: Ruta al PDF rotado
        """
        try:
            # Verificar que el ángulo de rotación es válido
            if rotation_angle not in [90, 180, 270]:
                raise ValueError("El ángulo de rotación debe ser 90, 180 o 270 grados")
            
            # Abrir el PDF
            with open(pdf_path, 'rb') as file:
                reader = PdfReader(file)
                writer = PdfWriter()
                
                # Rotar cada página
                for page in reader.pages:
                    page.rotate(rotation_angle)
                    writer.add_page(page)
                
                # Guardar el resultado
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                
                logger.info(f"PDF rotado creado exitosamente: {output_path}")
                return output_path
        
        except Exception as e:
            logger.error(f"Error al rotar PDF: {str(e)}")
            raise
    
    def cleanup(self):
        """Limpia recursos temporales."""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
                logger.debug(f"Directorio temporal eliminado: {self.temp_dir}")
            except Exception as e:
                logger.error(f"Error al eliminar directorio temporal: {str(e)}")
    
    def __del__(self):
        """Destructor que asegura la limpieza de recursos."""
        self.cleanup()