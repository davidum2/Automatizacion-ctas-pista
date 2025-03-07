from datetime import datetime
from babel.dates import format_date
import locale

def convert_fecha_to_texto(fecha_str):
    """
    Convierte una fecha en formato YYYY-MM-DD a texto en español.
    
    Args:
        fecha_str (str): Fecha en formato YYYY-MM-DD
        
    Returns:
        str: Fecha formateada en texto español
    """
    try:
        # Convertir la fecha a objeto datetime
        fecha_dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        
        # Establecer la configuración regional a español
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except:
            # Alternativa para Windows
            try:
                locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
            except:
                pass  # Si no se puede establecer, usar la configuración por defecto
        
        # Formatear la fecha en español con la primera letra en mayúscula
        fecha_formateada = format_date(fecha_dt, format="d 'de' MMMM 'del' y", locale='es')
        fecha_formateada = fecha_formateada[0].upper() + fecha_formateada[1:]
        
        return fecha_formateada
    except ValueError:
        raise ValueError("La fecha debe estar en formato YYYY-MM-DD")

def format_fecha_mensaje(fecha_str):
    """
    Formatea la fecha del mensaje en un formato especial.
    
    Args:
        fecha_str (str): Fecha en formato YYYY-MM-DD
        
    Returns:
        str: Fecha formateada para mensajes
    """
    try:
        # Convertir la fecha a objeto datetime
        fecha_dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        
        # Establecer la configuración regional a español
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except:
            # Alternativa para Windows
            try:
                locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
            except:
                pass
        
        # Formatear la fecha en formato especial
        fecha_formateada = format_date(fecha_dt, format="d MMM y", locale='es').upper()
        
        # Reemplazar abreviaturas de meses
        for mes_abr, mes_nuevo in [
            ('ENE', 'Ene.'), ('FEB', 'Feb.'), ('MAR', 'Mar.'), ('ABR', 'Abr.'),
            ('MAY', 'May.'), ('JUN', 'Jun.'), ('JUL', 'Jul.'), ('AGO', 'Ago.'),
            ('SEP', 'Sep.'), ('OCT', 'Oct.'), ('NOV', 'Nov.'), ('DIC', 'Dic.')
        ]:
            fecha_formateada = fecha_formateada.replace(mes_abr, mes_nuevo)
        
        return fecha_formateada
    except ValueError:
        raise ValueError("La fecha debe estar en formato YYYY-MM-DD")

def format_monto(monto):
    """
    Formatea un monto como moneda.
    
    Args:
        monto (float): Monto a formatear
        
    Returns:
        str: Monto formateado como moneda
    """
    return "$ {:,.2f}".format(monto)
