"""
Módulo para procesar PDFs de facturas y generar documentos combinados
"""
import os
import logging
import shutil
from utils.pdf_manager import PDFManager

# Configurar logging
logger = logging.getLogger(__name__)

class FacturaPDFProcessor:
    """
    Clase para procesar PDFs de facturas y generar documentos combinados.
    """
    
    def __init__(self, ui=None):
        """
        Inicializa el procesador de PDFs de facturas.
        
        Args:
            ui: Referencia opcional a la interfaz de usuario para reportes
        """
        self.ui = ui
        self.pdf_manager = PDFManager()
    
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
            logger.error(message)
        elif level == "warning":
            logger.warning(message)
        else:
            logger.info(message)
    
    def convert_word_documents(self, docx_files, output_dir):
        """
        Convierte documentos Word a PDF.
        
        Args:
            docx_files (dict): Diccionario con nombre:ruta de archivos DOCX
            output_dir (str): Directorio donde guardar los PDFs
            
        Returns:
            dict: Diccionario con rutas a los PDFs generados
        """
        self.update_status("Convirtiendo documentos Word a PDF...")
        pdf_files = {}
        
        # Procesar cada archivo DOCX
        for name, path in docx_files.items():
            try:
                if not os.path.exists(path):
                    self.update_status(f"Archivo no encontrado: {path}", "warning")
                    continue
                
                # Convertir a PDF
                pdf_path = self.pdf_manager.convert_docx_to_pdf(path, output_dir)
                pdf_files[name] = pdf_path
                
                self.update_status(f"PDF generado: {os.path.basename(pdf_path)}", "success")
                
            except Exception as e:
                self.update_status(f"Error al convertir {os.path.basename(path)}: {str(e)}", "error")
        
        return pdf_files
    
    def find_original_pdf(self, xml_path):
        """
        Busca el PDF original de la factura en la misma carpeta que el XML.
        
        Args:
            xml_path (str): Ruta al archivo XML
            
        Returns:
            str or None: Ruta al PDF encontrado o None si no se encuentra
        """
        try:
            # Obtener directorio y nombre base del XML
            xml_dir = os.path.dirname(xml_path)
            xml_base_name = os.path.splitext(os.path.basename(xml_path))[0]
            
            # Opciones para buscar el PDF original
            pdf_options = [
                # Mismo nombre que el XML
                os.path.join(xml_dir, f"{xml_base_name}.pdf"),
                # Nombre "factura.pdf"
                os.path.join(xml_dir, "factura.pdf"),
                # Nombre "Factura.pdf"
                os.path.join(xml_dir, "Factura.pdf"),
                # Cualquier archivo .pdf en el directorio
                *[os.path.join(xml_dir, f) for f in os.listdir(xml_dir) if f.lower().endswith('.pdf')]
            ]
            
            # Buscar el primer PDF que exista
            for pdf_path in pdf_options:
                if os.path.exists(pdf_path):
                    self.update_status(f"PDF original encontrado: {os.path.basename(pdf_path)}")
                    return pdf_path
            
            self.update_status("No se encontró el PDF original de la factura", "warning")
            return None
            
        except Exception as e:
            self.update_status(f"Error al buscar PDF original: {str(e)}", "error")
            return None
    
    def process_factura_pdfs(self, xml_path, output_dir, generated_docs):
        """
        Procesa los PDFs relacionados con una factura y genera un documento combinado.
        
        Args:
            xml_path (str): Ruta al archivo XML
            output_dir (str): Directorio de salida
            generated_docs (dict): Diccionario con documentos generados
            
        Returns:
            dict: Información sobre los PDFs procesados
        """
        try:
            self.update_status("Procesando PDFs de la factura...")
            
            # 1. Encontrar el PDF original de la factura
            factura_pdf_path = self.find_original_pdf(xml_path)
            if not factura_pdf_path:
                # Si no se encuentra el PDF original, no se puede continuar
                return None
            
            # 2. Convertir documentos Word a PDF
            docx_files = {
                name: path for name, path in generated_docs.items() 
                if path.lower().endswith('.docx')
            }
            
            pdf_files = self.convert_word_documents(docx_files, output_dir)
            
            # 3. Verificar que se tienen todos los PDFs necesarios
            required_pdfs = [
                'legalizacion_factura',
                'legalizacion_verificacion',
                'legalizacion_xml',
                'xml'
            ]
            
            missing_pdfs = [pdf for pdf in required_pdfs if pdf not in pdf_files]
            if missing_pdfs:
                self.update_status(f"Faltan PDFs necesarios: {', '.join(missing_pdfs)}", "warning")
                return None
            
            # 4. Buscar el PDF de verificación del SAT
            verificacion_sat_pdf = self.find_verificacion_sat_pdf(output_dir)
            if not verificacion_sat_pdf:
                # Crear un PDF vacío como sustituto
                verificacion_sat_pdf = self.create_empty_pdf(
                    os.path.join(output_dir, "verificacion_sat_empty.pdf"),
                    "No se encontró la verificación del SAT"
                )
            
            # 5. Crear documento combinado
            combined_pdf_path = os.path.join(output_dir, "documento_completo.pdf")
            
            result = self.pdf_manager.create_factura_legal_document(
                combined_pdf_path,
                factura_pdf_path,
                pdf_files['legalizacion_factura'],
                verificacion_sat_pdf,
                pdf_files['legalizacion_verificacion'],
                pdf_files['xml'],
                pdf_files['legalizacion_xml']
            )
            
            if result:
                self.update_status(f"Documento PDF combinado generado: {os.path.basename(combined_pdf_path)}", "success")
                
                # Devolver información sobre los PDFs procesados
                return {
                    'combined_pdf': combined_pdf_path,
                    'factura_pdf': factura_pdf_path,
                    'verificacion_sat_pdf': verificacion_sat_pdf,
                    'generated_pdfs': pdf_files
                }
            else:
                self.update_status("No se pudo generar el documento PDF combinado", "error")
                return None
            
        except Exception as e:
            self.update_status(f"Error al procesar PDFs de la factura: {str(e)}", "error")
            logger.exception("Error procesando PDFs de factura")
            return None
    
    def find_verificacion_sat_pdf(self, directory):
        """
        Busca el PDF de verificación del SAT en un directorio.
        
        Args:
            directory (str): Directorio donde buscar
            
        Returns:
            str or None: Ruta al PDF encontrado o None si no se encuentra
        """
        try:
            # Patrones comunes para archivos de verificación del SAT
            verification_patterns = [
                "Verificación de Comprobantes Fiscales Digitales por Internet",
                "verificacion_",
                "verificacion-",
                "sat_",
                "sat-"
            ]
            
            # Buscar archivos que coincidan con los patrones
            for file in os.listdir(directory):
                if file.lower().endswith('.pdf'):
                    for pattern in verification_patterns:
                        if pattern in file.lower():
                            pdf_path = os.path.join(directory, file)
                            self.update_status(f"Verificación SAT encontrada: {file}")
                            return pdf_path
            
            self.update_status("No se encontró la verificación del SAT", "warning")
            return None
            
        except Exception as e:
            self.update_status(f"Error al buscar verificación SAT: {str(e)}", "error")
            return None
    
    def create_empty_pdf(self, output_path, text="Documento no disponible"):
        """
        Crea un PDF con un mensaje simple.
        
        Args:
            output_path (str): Ruta donde guardar el PDF
            text (str): Texto a incluir en el PDF
            
        Returns:
            str: Ruta al PDF creado
        """
        try:
            from fpdf import FPDF
            
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt=text, ln=True, align='C')
            pdf.output(output_path)
            
            self.update_status(f"PDF vacío creado: {os.path.basename(output_path)}")
            return output_path
            
        except Exception as e:
            self.update_status(f"Error al crear PDF vacío: {str(e)}", "error")
            
            # En caso de error, intentar crear un PDF mínimo usando PyPDF2
            try:
                from PyPDF2 import PdfWriter
                
                writer = PdfWriter()
                writer.add_blank_page(width=612, height=792)  # Tamaño carta
                
                with open(output_path, 'wb') as f:
                    writer.write(f)
                
                return output_path
                
            except Exception as inner_e:
                self.update_status(f"Error secundario al crear PDF vacío: {str(inner_e)}", "error")
                return None