"""
Controlador para el procesamiento de facturas individuales
"""
import os
import logging
from datetime import datetime
from decimal import Decimal

# Importaciones internas
from core.xml_processor import XMLProcessor
from core.document_generator import DocumentGenerator
from utils.formatters import format_fecha_mensaje
from ui.dialogs import editar_conceptos

logger = logging.getLogger(__name__)

class FacturaController:
    """Controlador para el procesamiento de facturas individuales"""

    def __init__(self, ui):
        """
        Inicializa el controlador de facturas
        
        Args:
            ui: Referencia a la interfaz de usuario
        """
        self.ui = ui
        self.xml_processor = XMLProcessor()
        self.document_generator = DocumentGenerator(ui)
        
    def procesar_factura(self, xml_file, output_dir, partida, monto_formateado, datos_comunes):
        """
        Procesa una factura individual
        
        Args:
            xml_file: Ruta al archivo XML
            output_dir: Directorio de salida para los documentos
            partida: Información de la partida
            monto_formateado: Monto formateado de la partida
            datos_comunes: Datos comunes para el procesamiento
            
        Returns:
            dict: Información de la factura procesada o None si hay error
        """
        try:
            self.ui.update_status(f"🔍 Analizando XML: {os.path.basename(xml_file)}...")

            # 1. Extraer información base del XML
            xml_data = self.xml_processor.read_xml(xml_file)

            if not xml_data:
                self.ui.update_status(f"Error: No se pudo extraer información del XML", "error")
                return None

            # Extraer el monto de la factura correctamente
            if 'Total' in xml_data:
                try:
                    # Convertir el monto a Decimal para cálculos precisos
                    monto_decimal = Decimal(str(xml_data['Total']))
                    monto_formateado_factura = "$ {:,.2f}".format(monto_decimal)
                except:
                    # Si hay error, usar un valor por defecto
                    monto_decimal = Decimal('0.00')
                    monto_formateado_factura = "$ 0.00"
            else:
                monto_decimal = Decimal('0.00')
                monto_formateado_factura = "$ 0.00"

            # 2. Crear el diccionario data completo
            data = self._crear_diccionario_datos_completo(
                xml_data,
                partida,
                monto_formateado_factura,  # Usar el monto formateado de la factura
                datos_comunes
            )

            # Añadir explícitamente el monto decimal para cálculos posteriores
            data['monto_decimal'] = monto_decimal
            
            # Importante: Guardar la ruta del XML original para su uso posterior
            data['xml_path'] = xml_file

            # 3. Pre-procesar conceptos (formatearlos automáticamente)
            conceptos_str = self._formatear_conceptos_automatico(data['Conceptos'])

            # 4. Si está habilitado el editor de conceptos, mostrarlo
            from config import APP_CONFIG
            if APP_CONFIG.get('usar_editor_conceptos', True):
                self.ui.update_status(f"✏️ Abriendo editor de conceptos...")

                # Este es un punto crítico donde debemos esperar la interacción del usuario
                conceptos_editados = editar_conceptos(self.ui.root, data['Conceptos'], partida['descripcion'])

                if conceptos_editados:
                    data['Empleo_recurso'] = conceptos_editados
                else:
                    # Si no se editaron, usar los conceptos formateados automáticamente
                    data['Empleo_recurso'] = conceptos_str
            else:
                # Usar el formato automático
                data['Empleo_recurso'] = conceptos_str

            # 5. Generar documentos (DOCX y PDF)
            self.ui.update_status(f"📝 Generando documentos...")
            documento_results = self.document_generator.generate_all_documents(data, output_dir)

            # 6. Extraer rutas de documentos generados
            docx_files = documento_results.get('docx_files', {})
            pdf_files = {}
            pdf_combinado = None
            
            if documento_results.get('pdf_processed', False):
                pdf_data = documento_results.get('pdf_files', {})
                pdf_files = pdf_data.get('generated_pdfs', {})
                pdf_combinado = pdf_data.get('combined_pdf')

            # 7. Registro de éxito
            self.ui.update_status(
                f"✅ Documentos generados para factura {data['Serie']}{data['Numero']}",
                "success"
            )

            # 8. Retornar información para registro y relación
            return {
                'serie_numero': f"{data['Serie']}{data['Numero']}",
                'fecha': data.get('Fecha_factura_texto', data.get('Fecha_factura', '')),
                'fecha_factura': data.get('Fecha_factura', ''), 
                'emisor': data['Nombre_Emisor'],
                'rfc_emisor': data['Rfc_emisor'],
                'monto': data['monto'],
                'monto_decimal': data['monto_decimal'],  # Incluir el valor decimal para sumas posteriores
                'conceptos': data.get('Empleo_recurso', ''),
                'documentos': {
                    **docx_files,  # Documentos DOCX
                    'pdf_files': pdf_files,  # PDFs individuales
                    'pdf_combinado': pdf_combinado  # PDF combinado final
                }
            }

        except Exception as e:
            self.ui.update_status(
                f"Error al procesar factura {os.path.basename(xml_file)}: {str(e)}",
                "error"
            )
            logger.exception(f"Error procesando factura {xml_file}")
            return None
    
    def _crear_diccionario_datos_completo(self, xml_data, partida, monto_formateado, datos_comunes):
        """
        Crea un diccionario completo combinando todas las fuentes de datos
        
        Args:
            xml_data: Datos extraídos del XML
            partida: Información de la partida
            monto_formateado: Monto formateado para mostrar
            datos_comunes: Datos comunes del proceso
            
        Returns:
            dict: Diccionario completo de datos para generar documentos
        """
        # Crear un nuevo diccionario
        data = {}

        # 1. Agregar datos del XML
        data.update(xml_data)

        # 2. Agregar datos de la interfaz principal
        data['Fecha_doc'] = datos_comunes['fecha_documento_texto']
        data['Mes'] = datos_comunes['mes_asignado']

        # 3. Agregar información de la partida
        data['No_partida'] = partida['numero']
        data['Descripcion_partida'] = partida['descripcion']
        data['monto'] = monto_formateado

        # 4. Agregar información del personal seleccionado
        for key, value in datos_comunes['personal_recibio'].items():
            data[key] = value

        for key, value in datos_comunes['personal_vobo'].items():
            data[key] = value

        # 5. Información de fechas formateadas
        # Convertir ISO fecha a formato legible
        if 'Fecha_ISO' in data:
            fecha_obj = datetime.strptime(data['Fecha_ISO'].split('T')[0], '%Y-%m-%d')
            data['Fecha_original'] = data['Fecha_ISO']

            # Formato numérico para algunas ocasiones donde se necesite
            data['Fecha_factura'] = fecha_obj.strftime('%d/%m/%Y')

            # Formato textual en español
            import locale
            try:
                locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Intentar configurar locale español
            except:
                try:
                    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # Alternativa para Windows
                except:
                    pass  # Si no se puede establecer, usar la configuración por defecto

            # Usar format_date de babel para formatear con mes en texto
            from babel.dates import format_date
            data['Fecha_factura_texto'] = format_date(fecha_obj, format="d 'de' MMMM 'del' yyyy", locale='es')
            # Capitalizar primera letra
            if data['Fecha_factura_texto']:
                data['Fecha_factura_texto'] = data['Fecha_factura_texto'][0].upper() + data['Fecha_factura_texto'][1:]

        # 6. Información del Folio Fiscal
        if 'UUid' in data:
            data['Folio_Fiscal'] = data['UUid']

        # 7. Generar número de oficio
        data['No_of_remision'] = partida.get('numero_adicional', '')

        # 8. Auto-generar No_mensaje y Fecha_mensaje
        data['No_mensaje'] = partida.get('numero_adicional', '')
        data['Fecha_mensaje'] = format_fecha_mensaje(datos_comunes['fecha_documento'])

        return data
    
    def _formatear_conceptos_automatico(self, conceptos):
        """
        Formatea los conceptos para presentación automáticamente
        
        Args:
            conceptos: Diccionario de conceptos {descripcion: cantidad}
            
        Returns:
            str: Texto de conceptos formateado
        """
        from ui.dialogs import formatear_conceptos_automatico
        
        if not conceptos:
            return "Conceptos no disponibles"

        return formatear_conceptos_automatico(conceptos)