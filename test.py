#!/usr/bin/env python
"""
Script de datos simulados para la aplicación de automatización de documentos.
Este script NO modifica ningún archivo, solo define datos de ejemplo.
"""
import os
import logging
from decimal import Decimal
from datetime import datetime

# Configurar logging básico
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("datos-simulados")

def simular_datos_de_facturas():
    """
    Crea datos simulados de facturas para pruebas.
    """
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
    Crea datos comunes simulados para pruebas.
    """
    return {
        'excel_path': "./datos_simulados/excel_prueba.xlsx",
        'fecha_documento': "2025-03-13",
        'fecha_documento_texto': "13 de marzo del 2025",
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
        'base_dir': "./datos_simulados",
    }

def simular_partida():
    """
    Crea datos simulados de una partida para pruebas.
    """
    return {
        'numero': "24101",
        'descripcion': "Materiales y útiles de oficina",
        'monto': Decimal('1800.50'),
        'numero_adicional': "ABC/123/2025"
    }

def simular_xml_data_para_documento():
    """
    Crea datos simulados que se obtendrían de un XML para pruebas.
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

def simular_datos_factura_completos():
    """
    Simula un diccionario completo con todos los datos necesarios para generar documentos.
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
        'Fecha_factura_texto': '1 de marzo del 2025',
        'Fecha_doc': '13 de marzo del 2025',
        'Mes': 'marzo',
        'No_partida': '24101',
        'Descripcion_partida': 'Materiales y útiles de oficina',
        'monto': '$ 12,345.67',
        'Folio_Fiscal': '12345678-1234-1234-1234-123456789012',
        'No_mensaje': '123/2025',
        'Fecha_mensaje': '13 Mar. 2025',
        'Empleo_recurso': 'Adquisición de material de oficina',
        'Grado_recibio_la_compra': 'Cap. 1/o. Zpdrs., Enc. Tptes.',
        'Nombre_recibio_la_compra': 'Gustavo Trinidad Lizárraga Medrano.',
        'Matricula_recibio_la_compra': 'D-2432942',
        'Grado_Vo_Bo': 'Cor. Cab. E.M., Subjefe Admtvo.',
        'Nombre_Vo_Bo': 'Rafael López Rodríguez.',
        'Matricula_Vo_Bo': 'B-5767973',
        'No_of_remision': 'ABC/123/2025'
    }

def mostrar_ejemplo_uso():
    """
    Muestra ejemplos de cómo usar los datos simulados en tu código.
    """
    logger.info("=== EJEMPLOS DE USO DE DATOS SIMULADOS ===")

    # Ejemplo 1: Obtener facturas simuladas
    logger.info("\n--- Ejemplo 1: Facturas simuladas ---")
    facturas = simular_datos_de_facturas()
    logger.info(f"Se han generado {len(facturas)} facturas simuladas:")
    for i, factura in enumerate(facturas, 1):
        logger.info(f"Factura {i}: {factura['serie_numero']} - {factura['emisor']} - {factura['monto']}")

    # Ejemplo 2: Datos comunes
    logger.info("\n--- Ejemplo 2: Datos comunes ---")
    datos_comunes = simular_datos_comunes()
    logger.info(f"Fecha documento: {datos_comunes['fecha_documento_texto']}")
    logger.info(f"Mes asignado: {datos_comunes['mes_asignado']}")
    logger.info(f"Personal que recibe: {datos_comunes['personal_recibio']['Nombre_recibio_la_compra']}")

    # Ejemplo 3: Partidas
    logger.info("\n--- Ejemplo 3: Partida ---")
    partida = simular_partida()
    logger.info(f"Partida: {partida['numero']} - {partida['descripcion']}")
    logger.info(f"Monto asignado: ${partida['monto']}")

    # Ejemplo 4: Datos XML
    logger.info("\n--- Ejemplo 4: Datos XML ---")
    xml_data = simular_xml_data_para_documento()
    logger.info(f"Serie y Folio: {xml_data['Serie']}{xml_data['Numero']}")
    logger.info(f"Emisor: {xml_data['Emisor']['Nombre']} ({xml_data['Emisor']['Rfc']})")
    logger.info(f"Conceptos: {', '.join([f'{v} {k}' for k, v in xml_data['Conceptos'].items()])}")

    # Ejemplo 5: Datos completos de factura
    logger.info("\n--- Ejemplo 5: Datos completos ---")
    datos_completos = simular_datos_factura_completos()
    logger.info("Datos completos generados con todos los campos necesarios para plantillas")

    logger.info("\n=== FIN DE EJEMPLOS ===")
    logger.info("Puedes usar estos datos en tus pruebas importando las funciones de este módulo")

if __name__ == "__main__":
    logger.info("Script de datos simulados iniciado")
    logger.info("Este script NO modifica ningún archivo, solo define datos de ejemplo")
    mostrar_ejemplo_uso()
