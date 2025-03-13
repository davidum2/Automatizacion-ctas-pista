#!/usr/bin/env python
"""
Script de prueba para la generación de documentos de facturas
sin necesidad de iniciar todo el proceso
"""
import os
import sys
import logging
from decimal import Decimal
from datetime import datetime
from pathlib import Path

# Configurar logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("test-facturas")

# Asegurarse de que los módulos del proyecto estén disponibles
# Ajusta estas rutas según la estructura de tu proyecto
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.append(project_root)

# Importar las funciones necesarias desde tus módulos
try:
    # Importación del generador de documentos
    from generators.creacionDocumentos import creacionDocumentos
    from utils.formatters import convert_fecha_to_texto, format_fecha_mensaje, format_monto
    
    # Importar las funciones del nuevo código para procesar plantillas
    from generators.plantillas_partidas import procesar_plantillas_partida, calcular_montos_facturas
    
    # Verificar que los imports funcionan
    logger.info("Módulos importados correctamente")
except ImportError as e:
    logger.error(f"Error importando módulos: {e}")
    logger.error("Asegúrate de que las rutas a los módulos del proyecto sean correctas")
    sys.exit(1)

def simular_datos_de_facturas():
    """
    Crea datos simulados de facturas para probar la generación de documentos
    """
    # Datos simulados de facturas
    return [
        {
            'serie_numero': "A101",
            'fecha': "2025-03-01",
            'emisor': "Empresa Ejemplo 1, S.A. de C.V.",
            'rfc_emisor': "EEJ941231ABC",
            'monto': "$ 12,345.67",
            'conceptos': "10.000 Computadoras, 5.000 Monitores, 2.000 Teclados",
            'documentos': {
                'legalizacion_factura': 'ruta/simulada/legalizacion_factura.docx',
                'legalizacion_verificacion': 'ruta/simulada/legalizacion_verificacion.docx',
                'legalizacion_xml': 'ruta/simulada/legalizacion_xml.docx',
                'xml': 'ruta/simulada/xml.docx'
            }
        },
        {
            'serie_numero': "B202",
            'fecha': "2025-03-05",
            'emisor': "Distribuidora Ejemplo, S.A. de C.V.",
            'rfc_emisor': "DEJ960505XYZ",
            'monto': "$ 5,678.90",
            'conceptos': "20.000 Sillas, 10.000 Escritorios",
            'documentos': {
                'legalizacion_factura': 'ruta/simulada/legalizacion_factura.docx',
                'legalizacion_verificacion': 'ruta/simulada/legalizacion_verificacion.docx',
                'legalizacion_xml': 'ruta/simulada/legalizacion_xml.docx',
                'xml': 'ruta/simulada/xml.docx'
            }
        }
    ]

def simular_datos_comunes():
    """
    Crea datos comunes simulados para el procesamiento
    """
    # Obtener la fecha actual
    fecha_actual = datetime.now().strftime('%Y-%m-%d')
    
    # Convertir a formato texto en español
    fecha_texto = convert_fecha_to_texto(fecha_actual)
    
    return {
        'excel_path': "./test_output/excel_prueba.xlsx",
        'fecha_documento': fecha_actual,
        'fecha_documento_texto': fecha_texto,
        'mes_asignado': "marzo",
        'personal_recibio': {
            'Grado_recibio_la_compra': "Cap. 1/o. Zpdrs., Enc. Tptes.",
            'Nombre_recibio_la_compra': "Gustavo Trinidad Lizárraga Medrano.",
            'Matricula_recibio_la_compra': "D-2432942"
        },
        'personal_vobo': {
            'Grado_Vo_Bo': "Cor. Cab. E.M., Subjefe Admtvo.",
            'Nombre_Vo_Bo': "Rafael López Rodríguez.",
            'Matricula_Vo_Bo': "B-5767973"
        },
        'base_dir': "./test_output",
    }

def simular_partida():
    """
    Crea datos simulados de una partida para pruebas
    """
    return {
        'numero': "24101",
        'descripcion': "Materiales y útiles de oficina",
        'monto': Decimal('1800.50'),
        'numero_adicional': "ABC/123/2025"
    }

def crear_directorio_prueba():
    """
    Crea un directorio de prueba para los archivos generados
    """
    test_dir = os.path.join(current_dir, "test_output")
    os.makedirs(test_dir, exist_ok=True)
    return test_dir

def simular_xml_data_para_documento():
    """
    Crea datos simulados que se obtendrían de un XML para generar documentos
    """
    return {
        'xml': '<xml simulado>datos para representación</xml>',
        'Serie': 'A',
        'Numero': '101',
        'Fecha_ISO': '2025-03-01T12:00:00',
        'Total': '12345.67',
        'Emisor': {
            'Nombre': 'Empresa Ejemplo 1, S.A. de C.V.',
            'Rfc': 'EEJ941231ABC'
        },
        'Receptor': {
            'Nombre': 'Dependencia Gubernamental',
            'Rfc': 'DGU960101XYZ'
        },
        'Conceptos': {
            'Computadoras': 10,
            'Monitores': 5,
            'Teclados': 2
        },
        'Rfc_emisor': 'EEJ941231ABC',
        'Rfc_receptor': 'DGU960101XYZ',
        'UUid': '12345678-1234-1234-1234-123456789012',
        'Nombre_Emisor': 'Empresa Ejemplo 1, S.A. de C.V.',
        'Fecha_original': '2025-03-01T12:00:00',
        'Fecha_factura': '01/03/2025',
        'Fecha_factura_texto': '1 de marzo del 2025'
    }

def preparar_datos_para_documento(xml_data, partida, datos_comunes):
    """
    Crea un diccionario completo para generar documentos
    combinando datos del XML, partida y datos comunes
    """
    # Formato de monto
    monto_formateado = format_monto(float(xml_data['Total']))
    
    # Crear diccionario completo
    data = xml_data.copy()
    
    # Agregar datos adicionales
    data['Fecha_doc'] = datos_comunes['fecha_documento_texto']
    data['Mes'] = datos_comunes['mes_asignado']
    data['No_partida'] = partida['numero']
    data['Descripcion_partida'] = partida['descripcion']
    data['monto'] = monto_formateado
    data['Folio_Fiscal'] = data['UUid']
    data['No_mensaje'] = "123/2025"
    data['Fecha_mensaje'] = format_fecha_mensaje(datos_comunes['fecha_documento'])
    data['Empleo_recurso'] = "Adquisición de material de oficina"
    
    # Agregar datos del personal
    for key, value in datos_comunes['personal_recibio'].items():
        data[key] = value
    
    for key, value in datos_comunes['personal_vobo'].items():
        data[key] = value
    
    return data

def prueba_generar_documento():
    """
    Prueba la generación de un solo documento
    """
    # Dirección de prueba
    output_dir = crear_directorio_prueba()
    
    # Obtener datos simulados
    xml_data = simular_xml_data_para_documento()
    partida = simular_partida()
    datos_comunes = simular_datos_comunes()
    
    # Preparar datos para el documento
    data = preparar_datos_para_documento(xml_data, partida, datos_comunes)
    
    # Verificar si existe el directorio de plantillas
    templates_dir = os.path.join(project_root, "plantillas")
    if not os.path.exists(templates_dir):
        logger.warning(f"Directorio de plantillas no encontrado: {templates_dir}")
        logger.warning("Creando directorio de plantillas vacío para pruebas")
        os.makedirs(templates_dir, exist_ok=True)
    
    # Crear una plantilla de prueba si no existe
    test_template_path = os.path.join(templates_dir, "plantilla_prueba.docx")
    if not os.path.exists(test_template_path):
        logger.warning("No se encontró plantilla de prueba. Debes crear una manualmente.")
        logger.info(f"Ruta esperada de la plantilla: {test_template_path}")
        return
    
    try:
        # Generar documento de prueba
        logger.info("Generando documento de prueba...")
        output_path = creacionDocumentos(test_template_path, output_dir, data, "documento_prueba")
        
        logger.info(f"Documento generado exitosamente: {output_path}")
        logger.info(f"Revisa el directorio: {output_dir}")
        
    except Exception as e:
        logger.error(f"Error generando documento: {e}")

def prueba_procesar_plantillas_partida():
    """
    Prueba el procesamiento de plantillas para una partida
    """
    # Dirección de prueba
    output_dir = crear_directorio_prueba()
    
    # Obtener datos simulados
    facturas_info = simular_datos_de_facturas()
    partida = simular_partida()
    datos_comunes = simular_datos_comunes()
    
    try:
        # Imprimir la ubicación de las plantillas esperada
        expected_templates_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "plantillas")
        logger.info(f"Se buscarán plantillas en: {expected_templates_dir}")
        
        # Verificar si existen las plantillas necesarias
        required_templates = [
            "Ingresos y Egresos .xlsx",
            "Relacion Facturas.xlsx",
            "Oficio.docx"
        ]
        
        for template in required_templates:
            template_path = os.path.join(expected_templates_dir, template)
            if not os.path.exists(template_path):
                logger.warning(f"No se encontró la plantilla: {template}")
        
        # Calcular información resumida de facturas
        info_facturas = calcular_montos_facturas(facturas_info)
        logger.info(f"Calculados totales: {info_facturas['total_facturas']} facturas, "
                   f"monto total: {info_facturas['monto_formateado']}")
        
        # Añadir la información resumida a los datos comunes
        datos_comunes['info_facturas'] = info_facturas
        
        # Procesar las plantillas
        logger.info("Procesando plantillas de partida...")
        archivos_generados = procesar_plantillas_partida(partida, facturas_info, output_dir, datos_comunes)
        
        logger.info("Plantillas procesadas exitosamente:")
        for tipo, ruta in archivos_generados.items():
            logger.info(f"- {tipo}: {ruta}")
        
        logger.info(f"Revisa el directorio: {output_dir}")
        
    except Exception as e:
        logger.error(f"Error procesando plantillas: {e}")

def main():
    """Función principal para probar las funciones de generación de facturas"""
    logger.info("=== INICIANDO PRUEBA DE GENERACIÓN DE FACTURAS ===")
    
    # 1. Prueba básica de generación de un documento
    logger.info("\n--- Prueba de generación de documento individual ---")
    prueba_generar_documento()
    
    # 2. Prueba de procesamiento de plantillas para una partida
    logger.info("\n--- Prueba de procesamiento de plantillas de partida ---")
    prueba_procesar_plantillas_partida()
    
    logger.info("\n=== PRUEBA FINALIZADA ===")

if __name__ == "__main__":
    main()