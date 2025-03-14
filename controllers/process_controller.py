"""
Controlador principal del proceso de automatización
"""
import os
import time
import logging
from tkinter import messagebox
from datetime import datetime
from decimal import Decimal

# Importaciones internas
from utils.formatters import convert_fecha_to_texto
from core.excel_reader import ExcelReader
from controllers.partida_controller import PartidaController

logger = logging.getLogger(__name__)

class ProcessController:
    """Controlador del proceso principal de la aplicación"""

    def __init__(self, ui):
        """
        Inicializa el controlador del proceso
        
        Args:
            ui: Referencia a la interfaz de usuario
        """
        self.ui = ui
        
        # Componentes
        self.excel_reader = ExcelReader()
        self.partida_controller = PartidaController(ui)
        
        # Estadísticas de procesamiento
        self.facturas_procesadas = 0
        self.facturas_con_error = 0
        self.partidas_procesadas = 0
        
        # Variables para tiempo de procesamiento
        self.tiempo_inicio = None
        self.tiempos_operaciones = {}
        
    def iniciar_procesamiento(self, datos_interfaz):
        """
        Inicia el procesamiento a partir de los datos de la interfaz
        
        Args:
            datos_interfaz: Diccionario con los datos recopilados de la interfaz
        """
        # Reiniciar estadísticas
        self.facturas_procesadas = 0
        self.facturas_con_error = 0
        self.partidas_procesadas = 0
        
        # Reiniciar medición de tiempo
        self.medir_tiempo(None, True)
        
        try:
            # Completar datos comunes con información procesada
            datos_comunes = self._preparar_datos_comunes(datos_interfaz)
            
            # Procesar el archivo Excel
            self.ui.update_status("Leyendo archivo Excel de partidas...")
            partidas = self.excel_reader.read_partidas(datos_comunes['excel_path'])
            self.medir_tiempo("Lectura de Excel")
            
            self.ui.update_status(f"Se encontraron {len(partidas)} partidas en el archivo.", "success")
            
            # Procesar cada partida secuencialmente
            for i, partida in enumerate(partidas, 1):
                self.ui.update_status(f"\n--- Procesando partida {i}/{len(partidas)}: {partida['numero']} ---")
                self.ui.set_processing_state(True, f"Procesando partida {i}/{len(partidas)}...")
                
                # Verificar directorio de la partida
                partida_dir = os.path.join(datos_comunes['base_dir'], partida['numero'])
                if not os.path.exists(partida_dir):
                    self.ui.update_status(f"Directorio para partida {partida['numero']} no encontrado.", "warning")
                    continue
                
                # Procesar la partida
                resultado_partida = self.partida_controller.procesar_partida(
                    partida, partida_dir, datos_comunes
                )
                
                if resultado_partida:
                    self.partidas_procesadas += 1
                    self.facturas_procesadas += resultado_partida.get('facturas_procesadas', 0)
                    self.facturas_con_error += resultado_partida.get('facturas_con_error', 0)
                    
            # Proceso completado
            self._mostrar_resumen_final()
            
        except Exception as e:
            self.ui.update_status(f"Error general en el procesamiento: {str(e)}", "error")
            logger.exception("Error no controlado en el procesamiento")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
        finally:
            # Restaurar interfaz
            self.ui.set_processing_state(False)
    
    def _preparar_datos_comunes(self, datos_interfaz):
        """
        Prepara y completa los datos comunes para el procesamiento
        
        Args:
            datos_interfaz: Datos recopilados de la interfaz
            
        Returns:
            dict: Datos comunes completos para el procesamiento
        """
        # Convertir fecha a formato de texto
        try:
            fecha_documento_texto = convert_fecha_to_texto(datos_interfaz['fecha_documento'])
        except ValueError as e:
            self.ui.update_status(f"Error en formato de fecha: {str(e)}", "error")
            return datos_interfaz

        # Crear diccionario completo
        datos_comunes = {
            **datos_interfaz,  # Incluir todos los datos originales
            'fecha_documento_texto': fecha_documento_texto,
        }
        
        return datos_comunes
        
    def medir_tiempo(self, operacion, reiniciar=False):
        """
        Mide y registra el tiempo de operaciones
        
        Args:
            operacion: Nombre de la operación o None
            reiniciar: Si se debe reiniciar el contador
            
        Returns:
            float: Tiempo transcurrido
        """
        if reiniciar or self.tiempo_inicio is None:
            self.tiempo_inicio = time.time()
            self.tiempos_operaciones = {}
            return 0

        tiempo_actual = time.time()
        tiempo_transcurrido = tiempo_actual - self.tiempo_inicio

        if operacion:
            if operacion in self.tiempos_operaciones:
                self.tiempos_operaciones[operacion] += tiempo_transcurrido
            else:
                self.tiempos_operaciones[operacion] = tiempo_transcurrido

            self.ui.update_status(f"Operación '{operacion}' completada en {tiempo_transcurrido:.2f} segundos", "time")

        self.tiempo_inicio = tiempo_actual
        return tiempo_transcurrido
        
    def _mostrar_resumen_final(self):
        """Muestra el resumen final del procesamiento"""
        tiempo_total = sum(self.tiempos_operaciones.values())

        self.ui.update_status("\n===== RESUMEN DEL PROCESAMIENTO =====")
        self.ui.update_status(f"Total partidas procesadas: {self.partidas_procesadas}")
        self.ui.update_status(f"Total facturas procesadas: {self.facturas_procesadas}")
        self.ui.update_status(f"Facturas con error: {self.facturas_con_error}")
        self.ui.update_status(f"Tiempo total de procesamiento: {tiempo_total:.2f} segundos")

        # Mostrar tiempos por tipo de operación
        if self.tiempos_operaciones:
            self.ui.update_status("\nTiempos por operación:")
            for operacion, tiempo in sorted(self.tiempos_operaciones.items(), key=lambda x: x[1], reverse=True):
                if "completa" not in operacion:  # Excluir totales de partidas que ya están sumados en otros
                    porcentaje = (tiempo / tiempo_total) * 100
                    self.ui.update_status(f"  - {operacion}: {tiempo:.2f} segundos ({porcentaje:.1f}%)", "time")

        # Mensaje final
        mensaje_final = f"Proceso completado. {self.facturas_procesadas} facturas procesadas en {self.partidas_procesadas} partidas."
        self.ui.update_status(mensaje_final, "success")
        messagebox.showinfo("Proceso Completado", mensaje_final)