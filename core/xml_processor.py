import xml.etree.ElementTree as ET
from datetime import datetime
from babel.dates import format_date
import locale
import os

class XMLProcessor:
    """
    Clase para procesar archivos XML de facturas.
    """
    
    def read_xml(self, file_path, numero_mensaje, fecha_mensaje_asignacion, 
                mes_asignado, monto_asignado, fecha_documento, partida_numero, 
                no_of_remision, fecha_remision):
        """
        Lee y analiza un archivo XML para extraer información relevante.
        
        Args:
            file_path (str): Ruta al archivo XML
            numero_mensaje (str): Número de mensaje de asignación
            fecha_mensaje_asignacion (str): Fecha del mensaje en formato YYYY-MM-DD
            mes_asignado (str): Mes asignado
            monto_asignado (str): Monto asignado formateado
            fecha_documento (str): Fecha del documento formateada
            partida_numero (str): Número de partida
            no_of_remision (str): Número de oficio de remisión
            fecha_remision (str): Fecha de remisión formateada
            
        Returns:
            dict: Diccionario con la información extraída
        """
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            
            # Namespace
            ns = {
                'cfdi': 'http://www.sat.gob.mx/cfd/4',
                'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
            }
            
            # Buscar elementos específicos
            emisor = root.find('.//cfdi:Emisor', ns)
            receptor = root.find('.//cfdi:Receptor', ns)
            conceptos = root.find('.//cfdi:Conceptos', ns)
            complemento = root.find('.//cfdi:Complemento', ns)
            
            # Verificar si los elementos existen
            if emisor is None:
                raise ValueError("No se encontró la información del emisor en el XML")
            if receptor is None:
                raise ValueError("No se encontró la información del receptor en el XML")
            if conceptos is None:
                raise ValueError("No se encontraron conceptos en el XML")
            if complemento is None:
                raise ValueError("No se encontró el complemento en el XML")
            
            # Obtener el TimbreFiscalDigital
            complemento_info = complemento.find('.//tfd:TimbreFiscalDigital', ns)
            if complemento_info is None:
                raise ValueError("No se encontró el TimbreFiscalDigital en el XML")
                
            folio_fiscal = complemento_info.attrib['UUID']
            
            # Extraer información del emisor
            emisor_info = {
                'Nombre': emisor.attrib.get('Nombre', 'No especificado'),
                'Rfc': emisor.attrib.get('Rfc', 'No especificado'),
            }
            
            # Extraer información del receptor
            receptor_info = {
                'Nombre': receptor.attrib.get('Nombre', 'No especificado'),
                'Rfc': receptor.attrib.get('Rfc', 'No especificado'),
            }
            
            # Extraer y agrupar información de los conceptos
            conceptos_data = conceptos.findall('cfdi:Concepto', ns)
            agrupados = {}  # Diccionario para agrupar por descripción
            
            for concepto in conceptos_data:
                descripcion = concepto.attrib.get('Descripcion', '')
                cantidad = float(concepto.attrib.get('Cantidad', 0))
                
                if descripcion in agrupados:
                    agrupados[descripcion] += cantidad
                else:
                    agrupados[descripcion] = cantidad
            
            # Redondear las cantidades a 3 decimales
            for descripcion in agrupados:
                agrupados[descripcion] = round(agrupados[descripcion], 3)
            
            # Extraer y formatear la fecha
            fecha_original = root.attrib.get('Fecha', '')
            if not fecha_original:
                raise ValueError("No se encontró la fecha en el XML")
                
            fecha_formateada = datetime.strptime(fecha_original, '%Y-%m-%dT%H:%M:%S')
            
            # Establecer la configuración regional a español
            try:
                locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
            except:
                # Alternativa para Windows
                try:
                    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
                except:
                    pass  # Si no se puede establecer, usar la configuración por defecto
            
            # Formatear la fecha en español
            fecha_formateada = format_date(fecha_formateada, format="d 'de' MMMM 'del' y", locale='es')
            
            # Formatear la fecha del mensaje
            fecha_mensaje_dt = datetime.strptime(fecha_mensaje_asignacion, '%Y-%m-%d')
            fecha_mensaje_formateada = format_date(fecha_mensaje_dt, format="d MMM y", locale='es').upper()
            
            # Ajustar abreviaturas de meses
            for mes_abr, mes_nuevo in [
                ('ENE', 'Ene.'), ('FEB', 'Feb.'), ('MAR', 'Mar.'), ('ABR', 'Abr.'),
                ('MAY', 'May.'), ('JUN', 'Jun.'), ('JUL', 'Jul.'), ('AGO', 'Ago.'),
                ('SEP', 'Sep.'), ('OCT', 'Oct.'), ('NOV', 'Nov.'), ('DIC', 'Dic.')
            ]:
                fecha_mensaje_formateada = fecha_mensaje_formateada.replace(mes_abr, mes_nuevo)
            
            # Convertir todo el contenido del XML a una cadena de texto
            xml_string = ET.tostring(root, encoding='unicode')
            
            # Mapeo de partidas y sus descripciones
            partida_info = {
                '24101': {
                    'empleo_del_recurso': 'la adquisición de materiales y útiles de oficina',
                    'descripcion': 'MATERIALES Y ÚTILES DE OFICINA'
                },
                '31202': {
                    'empleo_del_recurso': 'la adquisición de Gas L.P.',
                    'descripcion': 'SERVICIO DE GAS'
                },
                '31401': {
                    'empleo_del_recurso': 'el pago del servicio telefónico',
                    'descripcion': 'SERVICIO TELÉFONO CONVENCIONAL'
                }
                # Agregar más partidas según sea necesario
            }
            
            # Obtener la información de la partida
            partida_data = partida_info.get(partida_numero, {
                'empleo_del_recurso': f'la partida {partida_numero}',
                'descripcion': f'PARTIDA {partida_numero}'
            })
            
            # Construir y devolver el diccionario de datos
            factura_data = {
                'No_mensaje': numero_mensaje,
                'Fecha_mensaje': fecha_mensaje_formateada,
                'Mes': mes_asignado,
                'monto': monto_asignado,
                'Fecha_doc': fecha_documento,
                'xml': xml_string,
                'Serie': root.attrib.get('Serie', ''),
                'Numero': root.attrib.get('Folio', ''),
                'Fecha_factura': fecha_formateada,
                'Fecha_original': fecha_original,
                'Total': root.attrib.get('Total', ''),
                'Emisor': emisor_info,
                'Receptor': receptor_info,
                'Conceptos': agrupados,
                'Rfc_emisor': emisor_info['Rfc'],
                'Rfc_receptor': receptor_info['Rfc'],
                'Folio_Fiscal': folio_fiscal,
                'No_partida': partida_numero,
                'Descripcion_partida': partida_data['descripcion'],
                'Empleo_recurso': partida_data['empleo_del_recurso'],
                'Nombre_Emisor': emisor_info['Nombre'],
                'No_of_remision': no_of_remision,
                'Fecha_remision': fecha_remision
            }
            
            return factura_data
            
        except Exception as e:
            raise Exception(f"Error al procesar el archivo XML: {str(e)}")
