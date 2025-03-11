import xml.etree.ElementTree as ET
from datetime import datetime
from babel.dates import format_date
import locale
import os

class XMLProcessor:
    """
    Clase para procesar archivos XML de facturas.
    """

    def read_xml(self, file_path):
        """
        Lee y analiza un archivo XML para extraer información relevante.

        Args:
            file_path (str): Ruta al archivo XML
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

            # Convertir todo el contenido del XML a una cadena de texto
            xml_string = ET.tostring(root, encoding='unicode')

            # Construir y devolver el diccionario de datos
            factura_data = {
                'xml': xml_string,
                'Serie': root.attrib.get('Serie', ''),
                'Numero': root.attrib.get('Folio', ''),
                'Fecha_ISO': fecha_original,
                'Total': root.attrib.get('Total', ''),
                'Emisor': emisor_info,
                'Receptor': receptor_info,
                'Conceptos': agrupados,  # Aquí usamos el diccionario agrupado
                'Rfc_emisor': emisor_info['Rfc'],
                'Rfc_receptor': receptor_info['Rfc'],
                'UUid': folio_fiscal,
                'Nombre_Emisor': emisor_info['Nombre'],
            }

            return factura_data

        except Exception as e:
            raise Exception(f"Error al procesar el archivo XML: {str(e)}")

