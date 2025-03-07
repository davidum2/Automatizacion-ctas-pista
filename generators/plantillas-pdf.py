import os
from fpdf import FPDF
import textwrap

class PDF(FPDF):
    """Clase personalizada para crear PDFs"""
    
    def header(self):
        """Define el encabezado que se repetirá en todas las páginas"""
        # Puedes personalizar el encabezado si es necesario
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'DOCUMENTO OFICIAL', 0, 1, 'C')
        self.ln(5)
    
    def footer(self):
        """Define el pie de página que se repetirá en todas las páginas"""
        # Puedes personalizar el pie de página si es necesario
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')


def createLegalizacionFactura(data, output_path):
    """
    Crea un PDF de legalización de factura.
    
    Args:
        data (dict): Datos de la factura
        output_path (str): Ruta de salida para el PDF
        
    Returns:
        str: Ruta al PDF generado
    """
    try:
        # Crear el objeto PDF
        pdf = PDF()
        pdf.add_page()
        
        # Configurar fuentes
        pdf.set_font('Arial', 'B', 16)
        
        # Título
        pdf.cell(0, 10, 'LEGALIZACIÓN DE FACTURA', 0, 1, 'C')
        pdf.ln(5)
        
        # Información
        pdf.set_font('Arial', '', 12)
        
        # Fecha
        pdf.cell(0, 10, f'Fecha: {data["Fecha_doc"]}', 0, 1)
        
        # Información de la factura
        pdf.cell(0, 10, f'Serie y Folio: {data["Serie"]}{data["Numero"]}', 0, 1)
        pdf.cell(0, 10, f'Fecha de Factura: {data["Fecha_factura"]}', 0, 1)
        pdf.cell(0, 10, f'Emisor: {data["Nombre_Emisor"]}', 0, 1)
        pdf.cell(0, 10, f'Monto: {data["monto"]}', 0, 1)
        
        # Texto de legalización
        pdf.ln(5)
        pdf.set_font('Arial', '', 11)
        
        texto = f"""
        Por medio de la presente se legaliza la factura electrónica No. {data['Serie']}{data['Numero']} de fecha {data['Fecha_factura']}, por la cantidad de {data['monto']} ({data['monto']} M.N.), por concepto de {data['Descripcion_partida']}, para {data['Empleo_recurso']}, con cargo a la partida {data['No_partida']}, correspondiente al mes de {data['Mes']} del año en curso, asignado con mensaje No. {data['No_mensaje']} de fecha {data['Fecha_mensaje']}.
        """
        
        # Ajustar el texto para que encaje en el ancho de la página
        texto = textwrap.fill(texto, width=80)
        
        # Agregar texto multilínea
        pdf.multi_cell(0, 10, texto)
        
        # Agregar espacio para firmas
        pdf.ln(20)
        pdf.cell(0, 10, 'Nombre y Firma del Responsable', 0, 1, 'C')
        
        # Guardar el PDF
        pdf.output(output_path)
        
        return output_path
    
    except Exception as e:
        raise Exception(f"Error al crear legalización de factura en PDF: {str(e)}")


def createLegalizacionVerificacionSAT(data, output_path):
    """
    Crea un PDF de legalización de verificación del SAT.
    
    Args:
        data (dict): Datos de la factura
        output_path (str): Ruta de salida para el PDF
        
    Returns:
        str: Ruta al PDF generado
    """
    try:
        # Crear el objeto PDF
        pdf = PDF()
        pdf.add_page()
        
        # Configurar fuentes
        pdf.set_font('Arial', 'B', 16)
        
        # Título
        pdf.cell(0, 10, 'LEGALIZACIÓN DE VERIFICACIÓN DEL SAT', 0, 1, 'C')
        pdf.ln(5)
        
        # Información
        pdf.set_font('Arial', '', 12)
        
        # Fecha
        pdf.cell(0, 10, f'Fecha: {data["Fecha_doc"]}', 0, 1)
        
        # Información de la factura
        pdf.cell(0, 10, f'Serie y Folio: {data["Serie"]}{data["Numero"]}', 0, 1)
        pdf.cell(0, 10, f'Fecha de Factura: {data["Fecha_factura"]}', 0, 1)
        pdf.cell(0, 10, f'Folio Fiscal (UUID): {data["Folio_Fiscal"]}', 0, 1)
        pdf.cell(0, 10, f'Emisor: {data["Nombre_Emisor"]}', 0, 1)
        pdf.cell(0, 10, f'Monto: {data["monto"]}', 0, 1)
        
        # Texto de legalización
        pdf.ln(5)
        pdf.set_font('Arial', '', 11)
        
        texto = f"""
        Por medio de la presente se legaliza la verificación del SAT correspondiente a la factura electrónica No. {data['Serie']}{data['Numero']} de fecha {data['Fecha_factura']}, por la cantidad de {data['monto']} ({data['monto']} M.N.), por concepto de {data['Descripcion_partida']}, para {data['Empleo_recurso']}, con cargo a la partida {data['No_partida']}, correspondiente al mes de {data['Mes']} del año en curso, asignado con mensaje No. {data['No_mensaje']} de fecha {data['Fecha_mensaje']}.
        """
        
        # Ajustar el texto para que encaje en el ancho de la página
        texto = textwrap.fill(texto, width=80)
        
        # Agregar texto multilínea
        pdf.multi_cell(0, 10, texto)
        
        # Agregar espacio para firmas
        pdf.ln(20)
        pdf.cell(0, 10, 'Nombre y Firma del Responsable', 0, 1, 'C')
        
        # Guardar el PDF
        pdf.output(output_path)
        
        return output_path
    
    except Exception as e:
        raise Exception(f"Error al crear legalización de verificación SAT en PDF: {str(e)}")


def cretaeLegalizacionXML(data, output_path):
    """
    Crea un PDF de legalización de XML.
    
    Args:
        data (dict): Datos de la factura
        output_path (str): Ruta de salida para el PDF
        
    Returns:
        str: Ruta al PDF generado
    """
    try:
        # Crear el objeto PDF
        pdf = PDF()
        pdf.add_page()
        
        # Configurar fuentes
        pdf.set_font('Arial', 'B', 16)
        
        # Título
        pdf.cell(0, 10, 'LEGALIZACIÓN DE XML', 0, 1, 'C')
        pdf.ln(5)
        
        # Información
        pdf.set_font('Arial', '', 12)
        
        # Fecha
        pdf.cell(0, 10, f'Fecha: {data["Fecha_doc"]}', 0, 1)
        
        # Información de la factura
        pdf.cell(0, 10, f'Serie y Folio: {data["Serie"]}{data["Numero"]}', 0, 1)
        pdf.cell(0, 10, f'Fecha de Factura: {data["Fecha_factura"]}', 0, 1)
        pdf.cell(0, 10, f'Folio Fiscal (UUID): {data["Folio_Fiscal"]}', 0, 1)
        pdf.cell(0, 10, f'Emisor: {data["Nombre_Emisor"]} (RFC: {data["Rfc_emisor"]})', 0, 1)
        pdf.cell(0, 10, f'Receptor: RFC: {data["Rfc_receptor"]}', 0, 1)
        pdf.cell(0, 10, f'Monto: {data["monto"]}', 0, 1)
        
        # Texto de legalización
        pdf.ln(5)
        pdf.set_font('Arial', '', 11)
        
        texto = f"""
        Por medio de la presente se legaliza el archivo XML correspondiente a la factura electrónica No. {data['Serie']}{data['Numero']} de fecha {data['Fecha_factura']}, por la cantidad de {data['monto']} ({data['monto']} M.N.), por concepto de {data['Descripcion_partida']}, para {data['Empleo_recurso']}, con cargo a la partida {data['No_partida']}, correspondiente al mes de {data['Mes']} del año en curso, asignado con mensaje No. {data['No_mensaje']} de fecha {data['Fecha_mensaje']}.
        """
        
        # Ajustar el texto para que encaje en el ancho de la página
        texto = textwrap.fill(texto, width=80)
        
        # Agregar texto multilínea
        pdf.multi_cell(0, 10, texto)
        
        # Agregar espacio para firmas
        pdf.ln(20)
        pdf.cell(0, 10, 'Nombre y Firma del Responsable', 0, 1, 'C')
        
        # Guardar el PDF
        pdf.output(output_path)
        
        return output_path
    
    except Exception as e:
        raise Exception(f"Error al crear legalización de XML en PDF: {str(e)}")


def createXMLenPDF(data, output_path):
    """
    Crea un PDF con el contenido del XML.
    
    Args:
        data (dict): Datos de la factura
        output_path (str): Ruta de salida para el PDF
        
    Returns:
        str: Ruta al PDF generado
    """
    try:
        # Crear el objeto PDF
        pdf = PDF()
        pdf.add_page()
        
        # Configurar fuentes
        pdf.set_font('Arial', 'B', 16)
        
        # Título
        pdf.cell(0, 10, 'COMPROBANTE FISCAL DIGITAL (CFDI)', 0, 1, 'C')
        pdf.ln(5)
        
        # Información básica
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, f'Folio Fiscal (UUID): {data["Folio_Fiscal"]}', 0, 1)
        pdf.cell(0, 10, f'Serie y Folio: {data["Serie"]}{data["Numero"]}', 0, 1)
        pdf.cell(0, 10, f'Fecha de Emisión: {data["Fecha_factura"]}', 0, 1)
        
        # Datos del emisor
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'DATOS DEL EMISOR:', 0, 1)
        
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f'Nombre: {data["Emisor"]["Nombre"]}', 0, 1)
        pdf.cell(0, 10, f'RFC: {data["Emisor"]["Rfc"]}', 0, 1)
        
        # Datos del receptor
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'DATOS DEL RECEPTOR:', 0, 1)
        
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f'Nombre: {data["Receptor"]["Nombre"]}', 0, 1)
        pdf.cell(0, 10, f'RFC: {data["Receptor"]["Rfc"]}', 0, 1)
        
        # Conceptos
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'CONCEPTOS:', 0, 1)
        
        # Configurar tabla
        pdf.set_font('Arial', 'B', 10)
        col_width = pdf.w / 2
        
        # Encabezados de tabla
        pdf.cell(col_width, 10, 'Descripción', 1, 0, 'C')
        pdf.cell(col_width, 10, 'Cantidad', 1, 1, 'C')
        
        # Filas de tabla
        pdf.set_font('Arial', '', 10)
        for descripcion, cantidad in data['Conceptos'].items():
            pdf.cell(col_width, 10, descripcion, 1, 0, 'L')
            pdf.cell(col_width, 10, str(cantidad), 1, 1, 'C')
        
        # Total
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, f'TOTAL: {data["monto"]}', 0, 1, 'R')
        
        # Nota
        pdf.ln(5)
        pdf.set_font('Arial', 'I', 10)
        texto = "NOTA: Este documento es una representación impresa del Comprobante Fiscal Digital (CFDI) y tiene validez fiscal."
        pdf.multi_cell(0, 10, texto)
        
        # Guardar el PDF
        pdf.output(output_path)
        
        return output_path
    
    except Exception as e:
        raise Exception(f"Error al crear XML en PDF: {str(e)}")
