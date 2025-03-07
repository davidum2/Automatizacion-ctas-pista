import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

def creacionDocumentos(template_path, output_dir, data, template_name):
    """
    Crea un documento de legalización de factura basado en una plantilla.
    
    Args:
        template_path (str): Ruta a la plantilla de legalización
        output_dir (str): Directorio donde se guardará el documento
        data (dict): Datos de la factura
        
    Returns:
        str: Ruta al documento generado
    """
    try:
        # Verificar que la plantilla existe
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"No se encontró la plantilla: {template_path}")
        
        # Cargar la plantilla
        doc = Document(template_path)
        
        def formatear_combustible(diccionario):
            """
            Identifica y formatea conceptos de combustible desde un diccionario de conceptos.
            
            Args:
                diccionario: Diccionario con conceptos de la factura
                
            Returns:
                str: Descripción formateada del combustible
            """
            # Palabras clave para identificar tipos de combustible
            diesel_keywords = ['diesel', 'diésel', 'gasóleo']
            gasolina_keywords = ['gasolina', 'premium', 'magna', 'regular', 'sin', 'verde', 'roja']
            
            litros_diesel = 0
            litros_gasolina = 0
            
            # Buscar en todas las claves del diccionario
            for concepto, cantidad in diccionario.items():
                concepto_lower = concepto.lower()
                
                # Verificar si es diesel
                if any(keyword in concepto_lower for keyword in diesel_keywords):
                    litros_diesel += cantidad
                
                # Verificar si es gasolina
                elif any(keyword in concepto_lower for keyword in gasolina_keywords):
                    litros_gasolina += cantidad
                
                # Si contiene la palabra "combustible" o "petrolífero" pero no se identificó tipo específico
                elif any(keyword in concepto_lower for keyword in ['combustible', 'petrolifero', 'petrolífero']):
                    # Verificar si hay indicios de diesel
                    if any(keyword in concepto_lower for keyword in ['pesado', 'alto azufre']):
                        litros_diesel += cantidad
                    else:
                        # Por defecto, asumimos que es gasolina
                        litros_gasolina += cantidad
            
            # Si se identificaron ambos tipos de combustible
            if litros_diesel > 0 and litros_gasolina > 0:
                return f"{litros_diesel:.3f} Lts. de diesel y {litros_gasolina:.3f} Lts. de gasolina"
            
            # Si solo se identificó diesel
            elif litros_diesel > 0:
                return f"{litros_diesel:.3f} Lts. de diesel"
            
            # Si solo se identificó gasolina
            elif litros_gasolina > 0:
                return f"{litros_gasolina:.3f} Lts. de gasolina"
            
            # Si no se identificó ningún combustible específico, buscar por unidades
            else:
                litros_total = 0
                for concepto, cantidad in diccionario.items():
                    concepto_lower = concepto.lower()
                    # Buscar indicios de que el concepto es medido en litros
                    if any(unit in concepto_lower for unit in ['lt', 'lts', 'litro', 'litros', '(lt)', '(l)']):
                        litros_total += cantidad
                
                if litros_total > 0:
                    return f"{litros_total:.3f} Lts. de combustible"
            
            # Si nada de lo anterior funciona, devolver un mensaje genérico o los conceptos originales
            conceptos_str = ", ".join([f"{concepto}: {cantidad}" for concepto, cantidad in diccionario.items()])
            if len(conceptos_str) > 100:  # Si es muy largo, cortarlo
                conceptos_str = conceptos_str[:97] + "..."
            
            return f"combustible ({conceptos_str})"   

        # Función para aplicar el formato de texto (Geomanist, 10pt)
        def aplicar_formato_texto(paragraph):
            for run in paragraph.runs:
                run.font.name = "Geomanist"
                run.font.size = Pt(10)

        # Buscar marcadores en el documento y reemplazarlos con los datos
        for paragraph in doc.paragraphs:
            # Reemplazar marcadores en el texto del párrafo
            if '{{XML}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{XML}}', data['xml'])
            if '{{FECHA_DOCUMENTO}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{FECHA_DOCUMENTO}}', data['Fecha_doc'])
            if '{{SERIE_NUMERO}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{SERIE_NUMERO}}', f"{data['Serie']}{data['Numero']}")
            if '{{FECHA_FACTURA}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{FECHA_FACTURA}}', data['Fecha_factura'])
            if '{{NOMBRE_EMISOR}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{NOMBRE_EMISOR}}', data['Nombre_Emisor'])
            if '{{MONTO}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{MONTO}}', data['monto'])
            if '{{EMPLEO_RECURSO}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{EMPLEO_RECURSO}}', formatear_combustible(data['Conceptos']))
            if '{{MES}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{MES}}', data['Mes'])
            if '{{NO_MENSAJE}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{NO_MENSAJE}}', data['No_mensaje'])
            if '{{FECHA_MENSAJE}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{FECHA_MENSAJE}}', data['Fecha_mensaje'])
            if '{{GRADO_RECIBIO_LA_COMPRA}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{GRADO_RECIBIO_LA_COMPRA}}', data['Grado_recibio_la_compra'])
            if '{{NOMBRE_RECIBIO_LA_COMPRA}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{NOMBRE_RECIBIO_LA_COMPRA}}', data['Nombre_recibio_la_compra'])
            if '{{MATRICULA_RECIBIO_LA_COMPRA}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{MATRICULA_RECIBIO_LA_COMPRA}}', data['Matricula_recibio_la_compra'])
            if '{{GRADO_VO_BO}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{GRADO_VO_BO}}', data['Grado_Vo_Bo'])
            if '{{NOMBRE_VO_BO}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{NOMBRE_VO_BO}}', data['Nombre_Vo_Bo'])
            if '{{MATRICULA_VO_BO}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{{MATRICULA_VO_BO}}', data['Matricula_Vo_Bo'])

            # Aplicar formato después de reemplazar el texto
            aplicar_formato_texto(paragraph)
        
        # También buscar en las tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if '{{FECHA_DOCUMENTO}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{FECHA_DOCUMENTO}}', data['Fecha_doc'])
                        if '{{SERIE_NUMERO}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{SERIE_NUMERO}}', f"{data['Serie']}{data['Numero']}")
                        if '{{FECHA_FACTURA}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{FECHA_FACTURA}}', data['Fecha_factura'])
                        if '{{NOMBRE_EMISOR}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{NOMBRE_EMISOR}}', data['Nombre_Emisor'])
                        if '{{MONTO}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{MONTO}}', data['monto'])
                        if '{{EMPLEO_RECURSO}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{EMPLEO_RECURSO}}', data['Empleo_recurso'])
                        if '{{GRADO_VO_BO}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{GRADO_VO_BO}}', data['Grado_Vo_Bo'])
                        if '{{NOMBRE_VO_BO}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{NOMBRE_VO_BO}}', data['Nombre_Vo_Bo'])
                        if '{{MATRICULA_VO_BO}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{MATRICULA_VO_BO}}', data['Matricula_Vo_Bo'])
                        if '{{MES}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{MES}}', data['Mes'])
                        if '{{NO_MENSAJE}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{NO_MENSAJE}}', data['No_mensaje'])
                        if '{{FECHA_MENSAJE}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{FECHA_MENSAJE}}', data['Fecha_mensaje'])
                        if '{{GRADO_RECIBIO_LA_COMPRA}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{GRADO_RECIBIO_LA_COMPRA}}', data['Grado_recibio_la_compra'])
                        if '{{NOMBRE_RECIBIO_LA_COMPRA}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{NOMBRE_RECIBIO_LA_COMPRA}}', data['Nombre_recibio_la_compra'])
                        if '{{MATRICULA_RECIBIO_LA_COMPRA}}' in paragraph.text:
                            paragraph.text = paragraph.text.replace('{{MATRICULA_RECIBIO_LA_COMPRA}}', data['Matricula_recibio_la_compra'])
                        
                        # Aplicar formato después de reemplazar el texto
                        aplicar_formato_texto(paragraph)

        # Guardar el documento
        output_path = os.path.join(output_dir, template_name + ".docx")
        doc.save(output_path)
        
        return output_path
        
    except Exception as e:
        raise Exception(f"Error al crear legalización de factura: {str(e)}")