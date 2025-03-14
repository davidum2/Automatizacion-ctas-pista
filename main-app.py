"""
Sistema de Automatización de Documentos por Partidas
Punto de entrada principal de la aplicación
"""
import sys
import os
import logging
import tkinter as tk

# Configurar el logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def setup_environment():
    """Configura el entorno para la aplicación"""
    # Añadir el directorio actual al path
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    
    # Verificar dependencias
    try:
        # Importar paquetes necesarios para verificar que estén instalados
        import pandas, docx, openpyxl, babel
        logger.info("Todas las dependencias están instaladas correctamente")
    except ImportError as e:
        logger.error(f"Error de dependencia: {e}")
        raise

def main():
    """Función principal de la aplicación"""
    try:
        # Configurar entorno
        setup_environment()
        
        # Importar componentes (después de setup para asegurar que el path esté correcto)
        from ui.app_window import AutomatizacionAppWindow
        
        # Iniciar interfaz gráfica
        root = tk.Tk()
        app = AutomatizacionAppWindow(root)
        root.mainloop()
        
    except Exception as e:
        logger.exception("Error al iniciar la aplicación")
        print(f"Error al iniciar la aplicación: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()