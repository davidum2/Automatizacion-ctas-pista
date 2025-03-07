import os
import docx2pdf

class FileUtils:
    """
    Clase de utilidades para el manejo de archivos.
    """
    
    def find_xml_files(self, base_dir):
        """
        Busca todos los archivos XML en una estructura de carpetas.
        
        Args:
            base_dir (str): Directorio base para iniciar la búsqueda
            
        Returns:
            list: Lista de rutas a archivos XML encontrados
        """
        xml_files = []
        
        # Recorrer todas las subcarpetas
        for root, _, files in os.walk(base_dir):
            for file in files:
                if file.lower().endswith('.xml'):
                    xml_files.append(os.path.join(root, file))
        
        return xml_files
    
    def find_subdirectories(self, base_dir):
        """
        Encuentra todos los subdirectorios directos de un directorio base.
        
        Args:
            base_dir (str): Directorio base para buscar subdirectorios
            
        Returns:
            list: Lista de rutas a subdirectorios
        """
        if not os.path.exists(base_dir):
            return []
        
        return [os.path.join(base_dir, d) for d in os.listdir(base_dir) 
                if os.path.isdir(os.path.join(base_dir, d))]
    
    def find_xml_in_directory(self, directory):
        """
        Busca archivos XML en un directorio específico (sin recursión).
        
        Args:
            directory (str): Directorio donde buscar
            
        Returns:
            list: Lista de rutas a archivos XML encontrados
        """
        if not os.path.exists(directory):
            return []
        
        return [os.path.join(directory, f) for f in os.listdir(directory) 
                if f.lower().endswith('.xml') and os.path.isfile(os.path.join(directory, f))]
    
    def find_pdf_for_xml(self, xml_file_path):
        """
        Busca un archivo PDF correspondiente a un archivo XML en la misma carpeta.
        
        Args:
            xml_file_path (str): Ruta al archivo XML
            
        Returns:
            str: Ruta al archivo PDF o None si no se encuentra
        """
        directory = os.path.dirname(xml_file_path)
        file_base = os.path.splitext(os.path.basename(xml_file_path))[0]
        
        # Primero, buscar un PDF con el mismo nombre base
        pdf_same_name = os.path.join(directory, file_base + '.pdf')
        if os.path.exists(pdf_same_name):
            return pdf_same_name
        
        # Si no se encuentra, buscar cualquier PDF en la misma carpeta
        for file in os.listdir(directory):
            if file.lower().endswith('.pdf'):
                return os.path.join(directory, file)
        
        return None


def convert_to_pdf(docx_path, output_folder):
    """
    Convierte un archivo DOCX a PDF.
    
    Args:
        docx_path (str): Ruta al archivo DOCX
        output_folder (str): Carpeta o ruta de archivo de salida
    """
    try:
        # Asegurar que Word se cierre correctamente
        docx2pdf.convert(
            docx_path,
            output_path=output_folder,
            keep_active=False,  # Forzar cierre de Word
        )
    except Exception as e:
        print(f"Error al convertir: {e}")
        raise
    finally:
        # Limpiar procesos residuales de Word
        try:
            import platform
            if platform.system() == 'Windows':
                os.system('taskkill /f /im winword.exe')
        except:
            pass  # Ignorar errores al intentar cerrar Word