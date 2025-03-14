"""
Configuración global de la aplicación
"""
import os

# Configuración de la aplicación
APP_CONFIG = {
    'usar_editor_conceptos': True,  # Activa o desactiva el editor de conceptos
    'formato_fecha': '%Y-%m-%d',    # Formato de fecha esperado en la interfaz
    'debug_mode': False,            # Modo de depuración
    'templates_dirs': [             # Directorios donde buscar plantillas
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "plantillas"),
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "templates"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantillas"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
    ]
}

# Información de personal predefinido
PERSONAL_RECIBE = [
    {
        'Grado_recibio_la_compra': "Cap. 1/o. Zpdrs., Enc. Tptes.",
        'Nombre_recibio_la_compra': "Gustavo Trinidad Lizárraga Medrano.",
        'Matricula_recibio_la_compra': "D-2432942"
    },
    {
        'Grado_recibio_la_compra': "Cor. Cab. E.M., Subjefe Admtvo.",
        'Nombre_recibio_la_compra': "Rafael López Rodríguez.",
        'Matricula_recibio_la_compra': "B-5767973"
    }
]

PERSONAL_VISTO_BUENO = [
    {
        'Grado_Vo_Bo': "Gral. Bgda. E.M., Cmte. C.N.A.",
        'Nombre_Vo_Bo': "Sergio Ángel Sánchez García.",
        'Matricula_Vo_Bo': "B-3628676"
    },
    {
        'Grado_Vo_Bo': "Garl. Brig. E.M., Jefe Edo. Myr.",
        'Nombre_Vo_Bo': "Samuel Javier Carreño.",
        'Matricula_Vo_Bo': "B-7094414"
    },
    {
        'Grado_Vo_Bo': "Cor. Cab. E.M., Subjefe Admtvo.",
        'Nombre_Vo_Bo': "Rafael López Rodríguez.",
        'Matricula_Vo_Bo': "B-5767973"
    }
]

# Lista de meses para la interfaz
MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]