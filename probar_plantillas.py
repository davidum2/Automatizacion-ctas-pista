#!/usr/bin/env python
"""
Script para probar las funciones de plantillas_partidas.py
con datos simulados y sin afectar archivos existentes.
"""
import os
import logging
from decimal import Decimal
from datetime import datetime

# Configurar logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("prueba-plantillas")

# Importar datos simulados
from test import (
    simular_datos_de_facturas,
    simular_datos_comunes,
    simular_partida
)

# Intentar importar las funciones de plantillas_partidas.py
try:
    from generators.plantillas_partidas import (
        procesar_plantillas_partida,
        procesar_plantilla_ingresos,
        procesar_plantilla_facturas,
        procesar_plantilla_oficio,
        calcular_montos_facturas
    )
    logger.info("Módulos de plantillas importados correctamente")
except ImportError as e:
    logger.error(f"Error importando funciones de plantillas: {e}")
    logger.error("Asegúrate de que la ruta al módulo generators.plantillas_partidas sea correcta")
    exit(1)

def crear_directorio_prueba():
    """
    Crea un directorio de prueba para los archivos generados,
    sin afectar ningún directorio existente.
    """
    # Crear un directorio de prueba único con timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    test_dir = os.path.join(os.getcwd(), f"prueba_plantillas_{timestamp}")
    os.makedirs(test_dir, exist_ok=True)
    logger.info(f"Creado directorio para pruebas: {test_dir}")
    return test_dir

def probar_procesar_plantillas_partida():
    """
    Prueba la función principal procesar_plantillas_partida
    con datos simulados.
    """
    # Crear directorio para pruebas
    output_dir = crear_directorio_prueba()

    # Obtener datos simulados
    facturas_info = simular_datos_de_facturas()
    partida = simular_partida()
    datos_comunes = simular_datos_comunes()

    try:
        # Calcular información resumida de facturas
        info_facturas = calcular_montos_facturas(facturas_info)
        logger.info(f"Calculados totales: {info_facturas['total_facturas']} facturas, "
                   f"monto total: {info_facturas['monto_formateado']}")

        # Añadir la información resumida a los datos comunes
        datos_comunes['info_facturas'] = info_facturas

        # Procesar plantillas
        logger.info(f"Llamando a procesar_plantillas_partida con directorio: {output_dir}")

        # Llamada a la función que queremos probar
        archivos_generados = procesar_plantillas_partida(partida, facturas_info, output_dir, datos_comunes)

        # Mostrar resultados
        logger.info("Plantillas procesadas exitosamente:")
        for tipo, ruta in archivos_generados.items():
            logger.info(f"- {tipo}: {ruta}")

        logger.info(f"Archivos generados en: {output_dir}")
        return True

    except Exception as e:
        logger.error(f"Error en prueba de procesar_plantillas_partida: {e}")
        import traceback
        traceback.print_exc()
        return False

def probar_plantilla_individual(tipo_plantilla):
    """
    Prueba una función de plantilla individual.

    Args:
        tipo_plantilla (str): 'ingresos', 'facturas' o 'oficio'
    """
    # Crear directorio para pruebas
    output_dir = crear_directorio_prueba()

    # Obtener datos simulados
    facturas_info = simular_datos_de_facturas()
    partida = simular_partida()
    datos_comunes = simular_datos_comunes()

    # Calcular información resumida de facturas
    info_facturas = calcular_montos_facturas(facturas_info)
    datos_comunes['info_facturas'] = info_facturas

    try:
        logger.info(f"Probando plantilla de {tipo_plantilla}...")

        if tipo_plantilla == 'ingresos':
            ruta = procesar_plantilla_ingresos(output_dir, partida, facturas_info, datos_comunes)
        elif tipo_plantilla == 'facturas':
            ruta = procesar_plantilla_facturas(output_dir, partida, facturas_info, datos_comunes)
        elif tipo_plantilla == 'oficio':
            ruta = procesar_plantilla_oficio(output_dir, partida, facturas_info, datos_comunes)
        else:
            logger.error(f"Tipo de plantilla desconocido: {tipo_plantilla}")
            return False

        logger.info(f"Plantilla de {tipo_plantilla} generada en: {ruta}")
        return True

    except Exception as e:
        logger.error(f"Error al procesar plantilla de {tipo_plantilla}: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Función principal para probar las funciones de plantillas_partidas.py"""
    logger.info("=== INICIANDO PRUEBAS DE PLANTILLAS_PARTIDAS.PY ===")

    # Probar todas las plantillas a la vez
    logger.info("\n--- Probando procesar_plantillas_partida (todas las plantillas) ---")
    if probar_procesar_plantillas_partida():
        logger.info("✅ Prueba de todas las plantillas completada con éxito")
    else:
        logger.error("❌ Error en la prueba de todas las plantillas")

    # Probar plantillas individuales
    for tipo in ['ingresos', 'facturas', 'oficio']:
        logger.info(f"\n--- Probando plantilla individual: {tipo} ---")
        if probar_plantilla_individual(tipo):
            logger.info(f"✅ Prueba de plantilla {tipo} completada con éxito")
        else:
            logger.error(f"❌ Error en la prueba de plantilla {tipo}")

    logger.info("\n=== PRUEBAS FINALIZADAS ===")

if __name__ == "__main__":
    main()
