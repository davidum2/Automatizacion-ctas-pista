import os
import re
from decimal import Decimal, InvalidOperation
import logging
from docx import Document
from datetime import datetime
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def calcular_montos_facturas(facturas_info):
    """
    Calcula información resumida de las facturas procesadas.

    Args:
        facturas_info (list): Lista de facturas procesadas

    Returns:
        dict: Diccionario con información resumida
    """
    # Inicializar contadores
    total_facturas = 0
    monto_total = Decimal('0.00')
    montos_individuales = []

    # Procesar cada factura
    for factura in facturas_info:
        # Saltear elementos que no son diccionarios
        if not isinstance(factura, dict):
            continue

        total_facturas += 1

        # Extraer el monto y convertirlo a Decimal
        monto_str = factura.get('monto', '0')
        if isinstance(monto_str, str):
            # Limpiar el string de monto (eliminar símbolos de moneda y separadores de miles)
            monto_limpio = re.sub(r'[^\d.]', '', monto_str.replace(',', ''))
            try:
                monto = Decimal(monto_limpio)
                montos_individuales.append(monto)
                monto_total += monto
            except InvalidOperation:
                logger.warning(f"No se pudo convertir el monto '{monto_str}' a decimal")
        elif isinstance(monto_str, (int, float)):
            monto = Decimal(str(monto_str))
            montos_individuales.append(monto)
            monto_total += monto

    # Formatear el monto total como string con formato moneda
    monto_formateado = f"$ {monto_total:,.2f}"

    return {
        'total_facturas': total_facturas,
        'monto_total': monto_total,
        'monto_formateado': monto_formateado,
        'montos_individuales': montos_individuales
    }

def aplicar_formato_geomanist(paragraph):
    """
    Aplica el formato Geomanist tamaño 10 a un párrafo.

    Args:
        paragraph: Párrafo al que aplicar el formato
    """
    for run in paragraph.runs:
        run.font.name = "Geomanist"
        run.font.size = Pt(10)

def aplicar_formato_a_documento(doc):
    """
    Aplica el formato Geomanist tamaño 10 a todo el documento.

    Args:
        doc: Documento Word a formatear
    """
    # Aplicar a párrafos
    for paragraph in doc.paragraphs:
        aplicar_formato_geomanist(paragraph)

    # Aplicar a tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    aplicar_formato_geomanist(paragraph)

def reemplazar_marcadores(doc, reemplazos):
    """
    Reemplaza marcadores en un documento preservando el formato.

    Args:
        doc: Documento Word
        reemplazos: Diccionario con los marcadores y sus reemplazos
    """
    # Reemplazar en párrafos
    for paragraph in doc.paragraphs:
        for key, value in reemplazos.items():
            for i, run in enumerate(paragraph.runs):
                run.text = run.text.replace(key, str(value))

    # Reemplazar en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in reemplazos.items():
                        for run in paragraph.runs:
                            run.text = run.text.replace(key, str(value))

def encontrar_plantilla(nombre_archivo, base_dir=None):
    """
    Busca una plantilla en diferentes ubicaciones posibles.

    Args:
        nombre_archivo: Nombre del archivo de plantilla
        base_dir: Directorio base opcional

    Returns:
        str: Ruta completa a la plantilla o None si no se encuentra
    """
    if not base_dir:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # Lista de posibles ubicaciones de plantillas
    directorios_posibles = [
        os.path.join(os.path.dirname(base_dir), "plantillas"),
        os.path.join(os.path.dirname(base_dir), "templates"),
        os.path.join(base_dir, "plantillas"),
        os.path.join(base_dir, "templates")
    ]

    # Buscar el archivo en cada directorio
    for directorio in directorios_posibles:
        ruta_posible = os.path.join(directorio, nombre_archivo)
        if os.path.exists(ruta_posible):
            return ruta_posible

    return None

def procesar_plantillas_partida(partida, facturas_info, partida_dir, datos_comunes):
    """
    Procesa todas las plantillas para una partida específica.
    """
    logger.info(f"Procesando plantillas para partida {partida.get('numero', 'desconocida')}")

    # Archivos generados
    archivos_generados = {}

    try:
        # Calcular información resumida de facturas
        info_facturas = calcular_montos_facturas(facturas_info)
        logger.info(f"Calculados totales para {info_facturas['total_facturas']} facturas. "
                   f"Monto total: {info_facturas['monto_formateado']}")

        # Añadir la información resumida a los datos comunes
        datos_comunes['info_facturas'] = info_facturas

        # Procesar plantilla de ingresos-egresos
        ruta_ingresos = procesar_plantilla_ingresos(
            partida_dir,
            partida,
            facturas_info,
            datos_comunes
        )
        archivos_generados["ingresos"] = ruta_ingresos
        logger.info(f"Plantilla de ingresos generada en: {ruta_ingresos}")

        # Procesar plantilla de relación de facturas
        ruta_facturas = procesar_plantilla_facturas(
            partida_dir,
            partida,
            facturas_info,
            datos_comunes
        )
        archivos_generados["facturas"] = ruta_facturas
        logger.info(f"Plantilla de facturas generada en: {ruta_facturas}")

        # Procesar plantilla de oficio
        ruta_oficio = procesar_plantilla_oficio(
            partida_dir,
            partida,
            facturas_info,
            datos_comunes
        )
        archivos_generados["oficio"] = ruta_oficio
        logger.info(f"Plantilla de oficio generada en: {ruta_oficio}")

        return archivos_generados

    except Exception as e:
        logger.error(f"Error al procesar plantillas de partida: {str(e)}")
        raise Exception(f"Error al procesar plantillas de partida: {str(e)}")

def procesar_plantilla_ingresos(output_dir, partida, facturas_info, datos_comunes):
    """
    Procesa la plantilla de ingresos/egresos en formato Word.
    """
    try:
        # Buscar la plantilla
        template_path = encontrar_plantilla("ingresos_egresos.docx")
        if not template_path:
            raise FileNotFoundError("No se encontró la plantilla de ingresos/egresos")

        logger.info(f"Utilizando plantilla: {template_path}")

        # Cargar la plantilla
        doc = Document(template_path)

        # Datos comunes
        mes = datos_comunes.get('mes_asignado', '').capitalize()
        partida_num = partida.get('numero', '')
        descripcion = partida.get('descripcion', '')
        fecha_doc = datos_comunes.get('fecha_documento_texto', '')

        # Información de facturas
        info_facturas = datos_comunes.get('info_facturas', {})
        total_facturas = info_facturas.get('total_facturas', 0)
        monto_total = info_facturas.get('monto_total', Decimal('0.00'))
        monto_formateado = info_facturas.get('monto_formateado', "$ 0.00")

        # Calcular montos específicos
        monto_asignado = Decimal(str(partida.get('monto', 0)))
        monto_asignado_str = f"$ {monto_asignado:,.2f}"
        aportacion = monto_total - monto_asignado
        aportacion_str = f"$ {aportacion:,.2f}"
        suma_ingresos = monto_asignado + aportacion
        suma_ingresos_str = f"$ {suma_ingresos:,.2f}"
        saldo = suma_ingresos - monto_total
        saldo_str = f"$ {saldo:,.2f}"

        # Datos del personal
        personal_vobo = datos_comunes.get('personal_vobo', {})
        grado_vobo = personal_vobo.get('Grado_Vo_Bo', '')
        nombre_vobo = personal_vobo.get('Nombre_Vo_Bo', '')
        matricula_vobo = personal_vobo.get('Matricula_Vo_Bo', '')

        # Crear diccionario de reemplazos
        reemplazos = {
            '{{FECHA_DOCUMENTO}}': fecha_doc,
            '{{MES}}': mes,
            '{{PARTIDA}}': partida_num,
            '{{DESCRIPCION}}': descripcion,
            '{{TOTAL_FACTURAS}}': str(total_facturas),
            '{{MONTO_TOTAL}}': monto_formateado,
            '{{MONTO}}': monto_asignado_str,
            '{{APORTACION}}': aportacion_str,
            '{{SUMA_INGRESOS}}': suma_ingresos_str,
            '{{EGRESOS}}': monto_formateado,
            '{{SALDO}}': saldo_str,
            '{{GRADO_VO_BO}}': grado_vobo,
            '{{NOMBRE_VO_BO}}': nombre_vobo,
            '{{MATRICULA_VO_BO}}': matricula_vobo
        }

        # Reemplazar todos los marcadores
        reemplazar_marcadores(doc, reemplazos)

        # Aplicar formato Geomanist 10pt a todo el documento
        aplicar_formato_a_documento(doc)

        # Guardar el documento
        output_path = os.path.join(output_dir, f"Ingresos_Egresos_Partida_{partida_num}.docx")
        doc.save(output_path)

        return output_path

    except Exception as e:
        logger.error(f"Error al procesar plantilla de ingresos: {str(e)}")
        raise Exception(f"Error al procesar plantilla de ingresos: {str(e)}")

# inicio relacion de facturas
# Importa lo necesario para trabajar con bordes y alineación
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Función para aplicar bordes a una celda
def aplicar_bordes_celda(celda):
    """
    Aplica bordes a todos los lados de una celda.

    Args:
        celda: Celda de tabla a la que aplicar bordes
    """
    # Código para agregar bordes a una celda
    tc = celda._tc
    tcPr = tc.get_or_add_tcPr()

    # Bordes: superior, inferior, izquierdo, derecho
    for border_type in ['top', 'bottom', 'left', 'right']:
        border = OxmlElement('w:{}Border'.format(border_type))
        border.set(qn('w:val'), 'single')  # Estilo de línea: single, double, dotted, etc.
        border.set(qn('w:sz'), '4')  # Ancho de línea (en octavos de punto)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')  # Color (RGB hex)

        borders = OxmlElement('w:borders')
        borders.append(border)

        tcBorders = tcPr.find_one(qn('w:tcBorders'))
        if tcBorders is None:
            tcBorders = OxmlElement('w:tcBorders')
            tcPr.append(tcBorders)

        tcBorder = tcBorders.find_one(qn('w:{}'.format(border_type)))
        if tcBorder is not None:
            tcBorders.remove(tcBorder)

        tcBorders.append(border)

# Función para aplicar formato a una celda (bordes, centrado, fuente)
def aplicar_formato_celda(celda, centrar=True, aplicar_bordes=True, fuente_geomanist=True):
    """
    Aplica formato completo a una celda: bordes, alineación y fuente.

    Args:
        celda: Celda de tabla
        centrar: Si el texto debe centrarse
        aplicar_bordes: Si se deben aplicar bordes
        fuente_geomanist: Si se debe aplicar fuente Geomanist
    """
    # Aplicar bordes
    if aplicar_bordes:
        aplicar_bordes_celda(celda)

    # Centrar texto en todos los párrafos de la celda
    for paragraph in celda.paragraphs:
        if centrar:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Aplicar formato Geomanist
        if fuente_geomanist:
            for run in paragraph.runs:
                run.font.name = "Geomanist"
                run.font.size = Pt(10)

# Modificación de tu función procesar_plantilla_facturas
def procesar_plantilla_facturas(output_dir, partida, facturas_info, datos_comunes):
    """
    Procesa la plantilla de relación de facturas en formato Word.
    """
    try:
        # [Código existente para cargar plantilla y preparar datos]

        # Verificar que hay al menos dos tablas en el documento
        if len(doc.tables) < 2:
            raise ValueError("El documento no contiene al menos dos tablas")

        tabla_facturas = doc.tables[1]
        logger.info(f"Se utilizará la segunda tabla con {len(tabla_facturas.rows)} filas y {len(tabla_facturas.columns)} columnas")

        # Filtrar facturas válidas
        facturas_validas = [f for f in facturas_info if isinstance(f, dict)]

        # Si hay más de una fila (encabezado + datos), eliminar todas excepto el encabezado
        while len(tabla_facturas.rows) > 1:
            tr = tabla_facturas._tbl.tr_lst.pop()
            tabla_facturas._tbl.remove(tr)

        # Asegurarse que la fila de encabezado tenga formato adecuado
        if len(tabla_facturas.rows) > 0:
            fila_encabezado = tabla_facturas.rows[0]
            for celda in fila_encabezado.cells:
                aplicar_formato_celda(celda, centrar=True, aplicar_bordes=True, fuente_geomanist=True)

                # Opcional: hacer el texto del encabezado en negrita
                for paragraph in celda.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

        # Ahora añadir una fila para cada factura
        for i, factura in enumerate(facturas_validas):
            # Añadir una nueva fila a la tabla
            nueva_fila = tabla_facturas.add_row()
            celdas = nueva_fila.cells

            # Verificar que la tabla tiene las columnas esperadas
            if len(celdas) >= 4:  # Fecha, Número, Emisor, Importe
                # Formatear fecha
                fecha_factura = factura.get('fecha', '')
                if isinstance(fecha_factura, str) and '-' in fecha_factura:
                    try:
                        fecha_obj = datetime.strptime(fecha_factura, '%Y-%m-%d')
                        fecha_factura = fecha_obj.strftime('%d/%m/%Y')
                    except:
                        pass

                # Asignar valores a las celdas
                celdas[0].text = str(fecha_factura)
                celdas[1].text = str(factura.get('serie_numero', ''))
                celdas[2].text = str(factura.get('emisor', ''))
                celdas[3].text = str(factura.get('monto', ''))

                # Aplicar formato a todas las celdas de la nueva fila
                for celda in celdas:
                    aplicar_formato_celda(celda, centrar=True, aplicar_bordes=True, fuente_geomanist=True)
            else:
                logger.warning(f"La tabla tiene {len(celdas)} columnas, menos de las 4 esperadas")

        # Añadir fila de total si hay facturas
        if facturas_validas:
            fila_total = tabla_facturas.add_row()
            celdas = fila_total.cells

            # Configurar celdas de total
            celdas[0].text = ""
            celdas[1].text = ""
            celdas[2].text = "TOTAL"
            celdas[3].text = monto_formateado

            # Aplicar formato a la fila de total
            # Para las celdas vacías
            aplicar_formato_celda(celdas[0], centrar=True, aplicar_bordes=True, fuente_geomanist=True)
            aplicar_formato_celda(celdas[1], centrar=True, aplicar_bordes=True, fuente_geomanist=True)

            # Para las celdas con "TOTAL" y el monto
            aplicar_formato_celda(celdas[2], centrar=False, aplicar_bordes=True, fuente_geomanist=True)
            celdas[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for run in celdas[2].paragraphs[0].runs:
                run.bold = True

            aplicar_formato_celda(celdas[3], centrar=False, aplicar_bordes=True, fuente_geomanist=True)
            celdas[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for run in celdas[3].paragraphs[0].runs:
                run.bold = True

        # Guardar el documento
        output_path = os.path.join(output_dir, f"Relacion_Facturas_Partida_{partida.get('numero', '')}.docx")
        doc.save(output_path)

        logger.info(f"Documento de relación de facturas generado: {output_path}")
        return output_path

    except Exception as e:
        logger.error(f"Error al procesar plantilla de facturas: {str(e)}")
        import traceback
        traceback.print_exc()
        raise Exception(f"Error al procesar plantilla de facturas: {str(e)}")



# fin de la relacion de facturas
def reemplazar_marcadores_preservando_formato(doc, reemplazos):
    """
    Reemplaza marcadores en un documento preservando totalmente el formato original.
    Solo reemplaza el texto dentro de los marcadores sin modificar ningún formato.

    Args:
        doc: Documento Word
        reemplazos: Diccionario con los marcadores y sus reemplazos
    """
    # Reemplazar en párrafos
    for paragraph in doc.paragraphs:
        texto_original = paragraph.text

        # Verificar si hay algún marcador en el párrafo
        tiene_marcador = any(key in texto_original for key in reemplazos)

        if tiene_marcador:
            # Guardar los runs originales y sus formatos
            runs_originales = []
            for run in paragraph.runs:
                runs_originales.append({
                    'texto': run.text,
                    'estilo': {
                        'bold': run.bold,
                        'italic': run.italic,
                        'underline': run.underline,
                        'font_name': run.font.name,
                        'font_size': run.font.size,
                        'color': run.font.color.rgb if run.font.color else None
                    }
                })

            # Realizar los reemplazos solo en el texto
            texto_nuevo = texto_original
            for key, value in reemplazos.items():
                if key in texto_nuevo:
                    texto_nuevo = texto_nuevo.replace(key, str(value))

            # Limpiar el párrafo original
            for _ in range(len(paragraph.runs)):
                paragraph.runs[0]._element.getparent().remove(paragraph.runs[0]._element)

            # Recrear los runs con el nuevo texto pero manteniendo el formato original
            index_texto_actual = 0
            for run_info in runs_originales:
                texto_run = run_info['texto']
                longitud_texto = len(texto_run)

                # Si el run original contenía parte del texto a reemplazar,
                # el texto nuevo puede ser más largo o más corto
                if index_texto_actual < len(texto_nuevo):
                    # Usar la misma longitud del run original o lo que queda del texto nuevo
                    texto_nuevo_run = texto_nuevo[index_texto_actual:min(index_texto_actual + longitud_texto, len(texto_nuevo))]

                    # Crear un nuevo run con el formato original
                    run = paragraph.add_run(texto_nuevo_run)

                    # Aplicar el formato original
                    estilo = run_info['estilo']
                    run.bold = estilo['bold']
                    run.italic = estilo['italic']
                    run.underline = estilo['underline']
                    if estilo['font_name']:
                        run.font.name = estilo['font_name']
                    if estilo['font_size']:
                        run.font.size = estilo['font_size']
                    if estilo['color']:
                        run.font.color.rgb = estilo['color']

                    # Avanzar en el índice del texto nuevo
                    index_texto_actual += longitud_texto

    # Reemplazar en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    texto_original = paragraph.text

                    # Verificar si hay algún marcador en el párrafo
                    tiene_marcador = any(key in texto_original for key in reemplazos)

                    if tiene_marcador:
                        # Guardar los runs originales y sus formatos
                        runs_originales = []
                        for run in paragraph.runs:
                            runs_originales.append({
                                'texto': run.text,
                                'estilo': {
                                    'bold': run.bold,
                                    'italic': run.italic,
                                    'underline': run.underline,
                                    'font_name': run.font.name,
                                    'font_size': run.font.size,
                                    'color': run.font.color.rgb if run.font.color else None
                                }
                            })

                        # Realizar los reemplazos solo en el texto
                        texto_nuevo = texto_original
                        for key, value in reemplazos.items():
                            if key in texto_nuevo:
                                texto_nuevo = texto_nuevo.replace(key, str(value))

                        # Limpiar el párrafo original
                        for _ in range(len(paragraph.runs)):
                            paragraph.runs[0]._element.getparent().remove(paragraph.runs[0]._element)

                        # Recrear los runs con el nuevo texto pero manteniendo el formato original
                        index_texto_actual = 0
                        for run_info in runs_originales:
                            texto_run = run_info['texto']
                            longitud_texto = len(texto_run)

                            if index_texto_actual < len(texto_nuevo):
                                texto_nuevo_run = texto_nuevo[index_texto_actual:min(index_texto_actual + longitud_texto, len(texto_nuevo))]

                                run = paragraph.add_run(texto_nuevo_run)

                                estilo = run_info['estilo']
                                run.bold = estilo['bold']
                                run.italic = estilo['italic']
                                run.underline = estilo['underline']
                                if estilo['font_name']:
                                    run.font.name = estilo['font_name']
                                if estilo['font_size']:
                                    run.font.size = estilo['font_size']
                                if estilo['color']:
                                    run.font.color.rgb = estilo['color']

                                index_texto_actual += longitud_texto

def procesar_plantilla_oficio(output_dir, partida, facturas_info, datos_comunes):
    """
    Procesa la plantilla de oficio en formato Word preservando formatos originales.
    Mantiene el tamaño de texto específico de algunos elementos e imágenes.
    """
    try:
        # Buscar la plantilla
        template_path = encontrar_plantilla("oficio.docx")
        if not template_path:
            raise FileNotFoundError("No se encontró la plantilla de oficio")

        logger.info(f"Utilizando plantilla: {template_path}")

        # Cargar la plantilla
        doc = Document(template_path)

        # Datos comunes
        mes = datos_comunes.get('mes_asignado', '').capitalize()
        partida_num = partida.get('numero', '')
        descripcion = partida.get('descripcion', '')
        fecha_doc = datos_comunes.get('fecha_documento_texto', '')

        # Información de facturas
        info_facturas = datos_comunes.get('info_facturas', {})
        total_facturas = info_facturas.get('total_facturas', 0)
        monto_formateado = info_facturas.get('monto_formateado', "$ 0.00")

        # Datos adicionales para el oficio
        no_oficio = partida.get('numero_adicional', '')

        # Datos del personal
        personal_vobo = datos_comunes.get('personal_vobo', {})
        grado_vobo = personal_vobo.get('Grado_Vo_Bo', '')
        nombre_vobo = personal_vobo.get('Nombre_Vo_Bo', '')
        matricula_vobo = personal_vobo.get('Matricula_Vo_Bo', '')

        # Crear diccionario de reemplazos
        reemplazos = {
            '{{FECHA_DOCUMENTO}}': fecha_doc,
            '{{MES}}': mes,
            '{{PARTIDA}}': partida_num,
            '{{DESCRIPCION}}': descripcion,
            '{{TOTAL_FACTURAS}}': str(total_facturas),
            '{{MONTO_TOTAL}}': monto_formateado,
            '{{NO_OFICIO}}': no_oficio,
            '{{GRADO_VO_BO}}': grado_vobo,
            '{{NOMBRE_VO_BO}}': nombre_vobo,
            '{{MATRICULA_VO_BO}}': matricula_vobo
        }

        # Reemplazar todos los marcadores preservando formatos originales
        reemplazar_marcadores_preservando_formato(doc, reemplazos)

        # No aplicamos formato Geomanist global en este caso para preservar los formatos originales

        # Guardar el documento
        output_path = os.path.join(output_dir, f"Oficio_Resumen_Partida_{partida_num}.docx")
        doc.save(output_path)

        logger.info(f"Documento de oficio generado: {output_path}")
        return output_path

    except Exception as e:
        logger.error(f"Error al procesar plantilla de oficio: {str(e)}")
        import traceback
        traceback.print_exc()
        raise Exception(f"Error al procesar plantilla de oficio: {str(e)}")
