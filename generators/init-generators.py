# Módulo generators
# Este módulo contiene las funciones para generar documentos específicos

from .oficio_de_remision import create_of_remision
from .legalizacionFactura import legalizacionFactura
from .legalizacionVerificacion import legalizacionVerificacion
from .legalizacionXml import legalizacionXml
from .crearDocXml import crearXML
from .createRelacionFacturas import create_relacion_de_facturas_excel

# Importar funciones de plantillas PDF
from .plantillas_pdf import (
    createLegalizacionFactura,
    createLegalizacionVerificacionSAT,
    cretaeLegalizacionXML,
    createXMLenPDF
)

__all__ = [
    'create_of_remision',
    'legalizacionFactura',
    'legalizacionVerificacion',
    'legalizacionXml',
    'crearXML',
    'create_relacion_de_facturas_excel',
    'createLegalizacionFactura',
    'createLegalizacionVerificacionSAT',
    'cretaeLegalizacionXML',
    'createXMLenPDF'
]
