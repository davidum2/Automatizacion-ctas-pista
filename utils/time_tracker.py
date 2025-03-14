"""
Utilidad para seguimiento de tiempo en operaciones
"""
import time
import logging

logger = logging.getLogger(__name__)

class TimeTracker:
    """Clase para realizar seguimiento del tiempo de operaciones"""
    
    def __init__(self, ui=None):
        """
        Inicializa el seguimiento de tiempo
        
        Args:
            ui: Referencia opcional a la interfaz de usuario para reportes
        """
        self.ui = ui
        self.tiempo_inicio = None
        self.tiempos_operaciones = {}
        self.reset()
        
    def reset(self):
        """Reinicia el seguimiento de tiempo"""
        self.tiempo_inicio = time.time()
        self.tiempos_operaciones = {}
        
    def measure(self, operacion=None):
        """
        Mide y registra el tiempo transcurrido para una operación
        
        Args:
            operacion: Nombre de la operación o None
            
        Returns:
            float: Tiempo transcurrido
        """
        tiempo_actual = time.time()
        
        if self.tiempo_inicio is None:
            self.tiempo_inicio = tiempo_actual
            return 0
            
        tiempo_transcurrido = tiempo_actual - self.tiempo_inicio

        if operacion:
            if operacion in self.tiempos_operaciones:
                self.tiempos_operaciones[operacion] += tiempo_transcurrido
            else:
                self.tiempos_operaciones[operacion] = tiempo_transcurrido

            # Reportar en la UI si está disponible
            if self.ui:
                self.ui.update_status(
                    f"Operación '{operacion}' completada en {tiempo_transcurrido:.2f} segundos", 
                    "time"
                )
            
            # Siempre registrar en el log
            logger.info(f"Operación '{operacion}' completada en {tiempo_transcurrido:.2f} segundos")

        # Actualizar el tiempo de inicio para la próxima medición
        self.tiempo_inicio = tiempo_actual
        return tiempo_transcurrido
        
    def get_summary(self):
        """
        Obtiene un resumen de los tiempos de operación
        
        Returns:
            dict: Resumen de tiempos
        """
        tiempo_total = sum(self.tiempos_operaciones.values())
        
        # Calcular porcentajes
        porcentajes = {}
        for operacion, tiempo in self.tiempos_operaciones.items():
            porcentajes[operacion] = (tiempo / tiempo_total * 100) if tiempo_total > 0 else 0
            
        return {
            'tiempo_total': tiempo_total,
            'tiempos': self.tiempos_operaciones,
            'porcentajes': porcentajes
        }
        
    def print_summary(self):
        """Imprime un resumen formateado de los tiempos"""
        summary = self.get_summary()
        
        logger.info("===== RESUMEN DE TIEMPOS =====")
        logger.info(f"Tiempo total: {summary['tiempo_total']:.2f} segundos")
        
        # Mostrar operaciones ordenadas por tiempo (de mayor a menor)
        for operacion, tiempo in sorted(
            summary['tiempos'].items(), 
            key=lambda x: x[1], 
            reverse=True
        ):
            porcentaje = summary['porcentajes'][operacion]
            logger.info(f"  - {operacion}: {tiempo:.2f} segundos ({porcentaje:.1f}%)")