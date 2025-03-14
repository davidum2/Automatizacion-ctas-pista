"""
Controlador para el procesamiento de partidas
"""
import os
import logging
from decimal import Decimal

# Importaciones internas
from controllers.factura_controller import FacturaController

logger = logging.getLogger(__name__)

class PartidaController:
    """Controlador para el procesamiento de partidas"""

    def __init__(self, ui):
        """
        Inicializa el controlador de partidas
        
        Args:
            ui: Referencia a la interfaz de usuario
        """
        self.ui = ui
        self.factura_controller = FacturaController(ui)
    
    def procesar_partida(self, partida, partida_dir, datos_comunes):
        """
        Procesa una partida y todas sus facturas
        
        Args:
            partida: Diccionario con información de la partida
            partida_dir: Directorio de la partida
            datos_comunes: Datos comunes para el procesamiento
            
        Returns:
            dict: Resultados del procesamiento de la partida o None si hay error
        """
        # Formatear monto de la partida
        monto_formateado = "$ {:,.2f}".format(partida['monto'])
        
        try:
            # Buscar facturas XML en la partida
            xml_files_in_partida = [
                f for f in os.listdir(partida_dir)
                if f.lower().endswith('.xml') and os.path.isfile(os.path.join(partida_dir, f))
            ]

            facturas_procesadas = 0
            facturas_con_error = 0
            facturas_info = []

            if xml_files_in_partida:
                # CASO 1: XML directamente en la carpeta de partida (una sola factura)
                self.ui.update_status(f"📄 Encontrado XML directamente en la carpeta de partida")

                # Procesar la factura única
                xml_file = os.path.join(partida_dir, xml_files_in_partida[0])
                resultado = self.factura_controller.procesar_factura(
                    xml_file, partida_dir, partida, monto_formateado, datos_comunes
                )

                if resultado:
                    facturas_procesadas += 1
                    facturas_info.append(resultado)
                else:
                    facturas_con_error += 1
            else:
                # CASO 2: Buscar en subcarpetas (múltiples facturas)
                subdirs = [
                    d for d in os.listdir(partida_dir)
                    if os.path.isdir(os.path.join(partida_dir, d))
                ]

                self.ui.update_status(f"📂 Partida {partida['numero']}: {len(subdirs)} subcarpetas encontradas.")

                # Procesar cada subcarpeta que contenga XML
                for subdir in subdirs:
                    factura_dir = os.path.join(partida_dir, subdir)
                    xml_files = [
                        f for f in os.listdir(factura_dir)
                        if f.lower().endswith('.xml') and os.path.isfile(os.path.join(factura_dir, f))
                    ]

                    if xml_files:
                        self.ui.update_status(f"  - Procesando factura en {subdir}...")

                        # Procesar la factura
                        xml_file = os.path.join(factura_dir, xml_files[0])
                        resultado = self.factura_controller.procesar_factura(
                            xml_file, factura_dir, partida, monto_formateado, datos_comunes
                        )

                        if resultado:
                            facturas_procesadas += 1
                            facturas_info.append(resultado)
                        else:
                            facturas_con_error += 1

            # Calcular el total de montos de las facturas
            monto_total = Decimal('0.00')
            for factura in facturas_info:
                if isinstance(factura, dict) and 'monto_decimal' in factura:
                    monto_total += factura['monto_decimal']
            
            # Formatear el monto total
            monto_total_formateado = "$ {:,.2f}".format(monto_total)
            
            # Añadir los datos de montos a los datos comunes para las plantillas
            datos_partida = {
                'monto_total': monto_total,
                'monto_total_formateado': monto_total_formateado
            }
            
            # Mostrar el total calculado
            self.ui.update_status(
                f"Monto total de facturas en partida {partida['numero']}: {monto_total_formateado}",
                "success"
            )

            # Generar relación de facturas si hay información disponible
            if facturas_info:
                self._generar_relacion_facturas(partida, facturas_info, partida_dir, datos_comunes, datos_partida)

            # Resumen de la partida
            self.ui.update_status(
                f"Partida {partida['numero']} completada: {facturas_procesadas} facturas procesadas, "
                f"{facturas_con_error} con errores.",
                "success" if facturas_con_error == 0 else "warning"
            )

            return {
                'numero': partida['numero'],
                'descripcion': partida['descripcion'],
                'facturas_procesadas': facturas_procesadas,
                'facturas_con_error': facturas_con_error,
                'monto_total': monto_total,
                'monto_total_formateado': monto_total_formateado
            }

        except Exception as e:
            self.ui.update_status(f"Error al procesar partida {partida['numero']}: {str(e)}", "error")
            logger.exception(f"Error procesando partida {partida['numero']}")
            return None
    
    def _generar_relacion_facturas(self, partida, facturas_info, partida_dir, datos_comunes, datos_partida=None):
        """
        Genera un documento de relación de facturas para la partida

        Args:
            partida: Datos de la partida
            facturas_info: Lista de información de facturas procesadas
            partida_dir: Directorio de la partida
            datos_comunes: Datos comunes para las plantillas
            datos_partida: Datos específicos de la partida, incluidos montos totales
        """
        try:
            self.ui.update_status(f"Generando relación de facturas para partida {partida['numero']}...")

            # Importar el módulo de plantillas de partidas
            from generators.plantillas_partidas import procesar_plantillas_partida, calcular_montos_facturas

            # Preparar datos comunes para las plantillas (hacer una copia para no modificar el original)
            datos_comunes_copia = datos_comunes.copy()
            
            # Si se proporcionaron datos específicos de la partida, utilizarlos
            if datos_partida and 'monto_total' in datos_partida:
                info_facturas = {
                    'total_facturas': len(facturas_info),
                    'monto_total': datos_partida['monto_total'],
                    'monto_formateado': datos_partida['monto_total_formateado'],
                    'montos_individuales': [f.get('monto_decimal', Decimal('0.00')) for f in facturas_info if isinstance(f, dict)]
                }
                self.ui.update_status(f"  - Usando monto total proporcionado: {datos_partida['monto_total_formateado']}")
            else:
                # Calcular información resumida de facturas (totales, montos, etc.)
                info_facturas = calcular_montos_facturas(facturas_info)
                self.ui.update_status(f"  - Calculados totales para {info_facturas['total_facturas']} facturas. "
                                f"Monto total: {info_facturas['monto_formateado']}")

            # Añadir la información resumida a los datos comunes
            datos_comunes_copia['info_facturas'] = info_facturas

            # Procesar todas las plantillas de la partida
            self.ui.update_status("Procesando plantillas de documentos...")
            archivos_generados = procesar_plantillas_partida(
                partida,
                facturas_info,
                partida_dir,
                datos_comunes_copia
            )

            # Registrar los archivos generados
            if archivos_generados:
                for tipo, ruta in archivos_generados.items():
                    self.ui.update_status(f"  - {tipo.capitalize()}: {os.path.basename(ruta)}")

            self.ui.update_status(
                f"✅ Relación de facturas generada para partida {partida['numero']}",
                "success"
            )

            return archivos_generados
        except Exception as e:
            self.ui.update_status(
                f"Error al generar relación de facturas: {str(e)}",
                "error"
            )
            logger.exception(f"Error al generar relación de facturas para {partida['numero']}")
            return None