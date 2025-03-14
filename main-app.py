"""
Sistema de Automatización de Documentos por Partidas
Implementación secuencial con UI responsiva
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import time
import locale
import logging
from babel.dates import format_date

# Importar módulos refactorizados
from core.excel_reader import ExcelReader
from core.xml_processor import XMLProcessor
from core.document_generator import DocumentGenerator
from utils.file_utils import FileUtils, convert_to_pdf
from ui.date_selector import DateSelector
from utils.formatters import convert_fecha_to_texto, format_fecha_mensaje
# Implementación directa para editar conceptos (evitar problemas de importación)
import re

def formatear_conceptos_automatico(conceptos_originales):
    """
    Formatea automáticamente los conceptos sin necesidad de interfaz gráfica

    Args:
        conceptos_originales (dict): Diccionario con los conceptos originales {descripcion: cantidad}

    Returns:
        str: Texto formateado de conceptos
    """
    total_items = sum(conceptos_originales.values())

    # Si solo hay un concepto, usar ese directamente
    if len(conceptos_originales) == 1:
        descripcion = list(conceptos_originales.keys())[0]
        cantidad = list(conceptos_originales.values())[0]
        return f"{cantidad:.3f} {descripcion}"

    # Si hay 2-3 conceptos, listarlos todos
    elif len(conceptos_originales) <= 3:
        conceptos_texto = []
        for descripcion, cantidad in conceptos_originales.items():
            # Limpiar descripción
            clean_desc = re.sub(r'^\d+\s*\.\s*', '', descripcion).strip()
            conceptos_texto.append(f"{cantidad:.3f} {clean_desc}")
        return ", ".join(conceptos_texto)

    # Si hay muchos conceptos, hacer un resumen
    else:
        # Tomar los 3 conceptos más importantes
        sorted_items = sorted(conceptos_originales.items(), key=lambda x: x[1], reverse=True)
        principales = sorted_items[:3]

        conceptos_texto = []
        for descripcion, cantidad in principales:
            clean_desc = re.sub(r'^\d+\s*\.\s*', '', descripcion).strip()
            conceptos_texto.append(f"{cantidad:.3f} {clean_desc}")

        return f"{', '.join(conceptos_texto)} y otros artículos (total {total_items:.3f} unidades)"


class SimpleConceptoDialog(tk.simpledialog.Dialog):
    """Un diálogo simple y ligero para editar conceptos"""

    def __init__(self, parent, conceptos_originales, partida_descripcion):
        self.conceptos_originales = conceptos_originales
        self.partida_descripcion = partida_descripcion
        self.sugerencia = formatear_conceptos_automatico(conceptos_originales)

        # Título corto para evitar problemas de ancho
        super().__init__(parent, title="Editar Conceptos")

    def body(self, master):
        """Crear el cuerpo del diálogo"""
        # Frame principal con padding mínimo
        frame = tk.Frame(master)
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Información de la partida (reducida al mínimo)
        tk.Label(frame, text=f"Partida: {self.partida_descripcion[:50]}...",
                anchor='w').pack(fill='x')

        # Espacio mínimo
        tk.Frame(frame, height=5).pack()

        # Texto simple de instrucción
        tk.Label(frame, text="Edite el texto de conceptos:").pack(anchor='w')

        # Campo de texto
        self.texto_conceptos = tk.Text(frame, height=10, width=60, wrap='word')
        self.texto_conceptos.pack(fill='both', expand=True)
        self.texto_conceptos.insert('1.0', self.sugerencia)

        # Sin scrollbar para reducir complejidad

        # Botón para restaurar sugerencia
        tk.Button(frame, text="Restaurar Sugerencia",
                 command=self.restaurar_sugerencia).pack(anchor='w')

        # Para dar foco al campo de texto
        self.texto_conceptos.focus_set()
        return frame

    def buttonbox(self):
        """Personalizar los botones para que sean más simples"""
        box = tk.Frame(self)

        # Botones simplificados
        w = tk.Button(box, text="Aceptar", width=10, command=self.ok)
        w.pack(side='left', padx=5, pady=5)
        w = tk.Button(box, text="Cancelar", width=10, command=self.cancel)
        w.pack(side='left', padx=5, pady=5)

        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)

        box.pack()

    def restaurar_sugerencia(self):
        """Restaura la sugerencia en el campo de texto"""
        self.texto_conceptos.delete('1.0', 'end')
        self.texto_conceptos.insert('1.0', self.sugerencia)

    def validate(self):
        """Validar la entrada"""
        self.result = self.texto_conceptos.get('1.0', 'end-1c').strip()
        if not self.result:
            tk.messagebox.showwarning("Advertencia", "El texto no puede estar vacío")
            return False
        return True

    def apply(self):
        """El resultado ya se guardó en validate()"""
        pass


def editar_conceptos_simple(parent, conceptos_originales, partida_descripcion):
    """
    Muestra un diálogo simple para editar conceptos y devuelve el texto editado.
    Si el usuario cancela, devuelve la sugerencia automática.

    Args:
        parent: Ventana padre
        conceptos_originales (dict): Diccionario con los conceptos originales
        partida_descripcion (str): Descripción de la partida

    Returns:
        str: Texto de conceptos editado o sugerencia automática
    """
    try:
        # Generar una sugerencia automática primero
        sugerencia = formatear_conceptos_automatico(conceptos_originales)

        # Mostrar el diálogo simplificado
        dialog = SimpleConceptoDialog(parent, conceptos_originales, partida_descripcion)

        # Si se canceló o dio error, usar la sugerencia
        if not hasattr(dialog, 'result') or not dialog.result:
            return sugerencia

        return dialog.result

    except Exception as e:
        # Si hay cualquier error, devolver la sugerencia automática
        print(f"Error en editor de conceptos: {e}")
        return sugerencia

# Configuración de logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ===================================================
# APLICACIÓN PRINCIPAL
# ===================================================

class AutomatizacionApp:
    """Interfaz principal de la aplicación"""

    def __init__(self, root):
        self.root = root
        self.root.title("Automatización de Documentos por Partidas")
        self.root.geometry("800x700")

        # Configuración
        self.config = {
            'usar_editor_conceptos': True,  # Activa o desactiva el editor de conceptos
        }

        # Inicializar componentes
        self.excel_reader = ExcelReader()
        self.xml_processor = XMLProcessor()
        self.document_generator = DocumentGenerator()
        self.file_utils = FileUtils()

        # Estadísticas de procesamiento
        self.facturas_procesadas = 0
        self.facturas_con_error = 0
        self.partidas_procesadas = 0

        # Variables para tiempo de procesamiento
        self.tiempo_inicio = None
        self.tiempos_operaciones = {}

        # Lista de personal predefinido
        self.lista_personal_recibe = self._cargar_personal_recibe()
        self.lista_personal_visto_bueno = self._cargar_personal_visto_bueno()

        # Crear la interfaz
        self.create_widgets()

    def _cargar_personal_recibe(self):
        """Carga lista predefinida de personal que recibe compras"""
        return [
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

    def _cargar_personal_visto_bueno(self):
        """Carga lista predefinida de personal que da visto bueno"""
        return [
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

    def create_widgets(self):
        """Crea los componentes de la interfaz gráfica"""
        # Configurar grid
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, minsize=120)

        # Selección de archivo Excel
        tk.Label(self.root, text="Archivo Excel de Partidas:", anchor='w').grid(
            row=0, column=0, padx=10, pady=5, sticky='ew')
        self.entry_excel_path = tk.Entry(self.root, width=50)
        self.entry_excel_path.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=self.select_excel_file).grid(
            row=0, column=2, padx=10, pady=5)

        # Separador
        ttk.Separator(self.root, orient='horizontal').grid(
            row=1, column=0, columnspan=3, sticky='ew', pady=10)

        # Sección de información común
        tk.Label(self.root, text="Información Común", font=('Helvetica', 10, 'bold')).grid(
            row=2, column=0, columnspan=3, sticky='w', padx=5)

        # Fecha del documento
        tk.Label(self.root, text="Fecha de elaboración del documento:", anchor='w').grid(
            row=3, column=0, padx=10, pady=5, sticky='ew')
        self.entry_fecha_documento = tk.Entry(self.root)
        self.entry_fecha_documento.grid(row=3, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar",
                  command=lambda: self.select_date(self.entry_fecha_documento)).grid(
            row=3, column=2, padx=10, pady=5)

        # Mes asignado
        tk.Label(self.root, text="Mes asignado:", anchor='w').grid(
            row=4, column=0, padx=10, pady=5, sticky='ew')
        self.mes_asignado_var = tk.StringVar(self.root)
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        self.mes_asignado_var.set(meses[datetime.now().month - 1])  # Mes actual predeterminado
        option_menu_meses = tk.OptionMenu(self.root, self.mes_asignado_var, *meses)
        option_menu_meses.grid(row=4, column=1, padx=10, pady=5, sticky='ew')

        # Separador para sección de personal
        ttk.Separator(self.root, orient='horizontal').grid(
            row=5, column=0, columnspan=3, sticky='ew', pady=10)

        # Sección de Personal
        tk.Label(self.root, text="Información de Personal", font=('Helvetica', 10, 'bold')).grid(
            row=6, column=0, columnspan=3, sticky='w', padx=5)

        # Persona que recibió la compra
        tk.Label(self.root, text="Persona que recibió la compra:", anchor='w').grid(
            row=7, column=0, padx=10, pady=5, sticky='ew')

        # Crear combobox para personal que recibió
        self.personal_recibio_var = tk.StringVar(self.root)
        opciones_recibio = self._generar_opciones_personal(self.lista_personal_recibe,
                                                         ['Grado_recibio_la_compra',
                                                          'Nombre_recibio_la_compra',
                                                          'Matricula_recibio_la_compra'])
        self.combo_personal_recibio = ttk.Combobox(self.root,
                                                 textvariable=self.personal_recibio_var,
                                                 values=opciones_recibio,
                                                 width=60)
        self.combo_personal_recibio.grid(row=7, column=1, padx=10, pady=5, sticky='ew')

        # Persona que dio el Vo. Bo.
        tk.Label(self.root, text="Persona que dio el Vo. Bo.:", anchor='w').grid(
            row=8, column=0, padx=10, pady=5, sticky='ew')
        self.personal_vobo_var = tk.StringVar(self.root)
        opciones_vobo = self._generar_opciones_personal(self.lista_personal_visto_bueno,
                                                      ['Grado_Vo_Bo',
                                                       'Nombre_Vo_Bo',
                                                       'Matricula_Vo_Bo'])
        self.combo_personal_vobo = ttk.Combobox(self.root,
                                              textvariable=self.personal_vobo_var,
                                              values=opciones_vobo,
                                              width=60)
        self.combo_personal_vobo.grid(row=8, column=1, padx=10, pady=5, sticky='ew')

        # Separador para indicador de estado
        ttk.Separator(self.root, orient='horizontal').grid(
            row=9, column=0, columnspan=3, sticky='ew', pady=10)

        # Etiqueta de estado de procesamiento
        self.state_label_var = tk.StringVar(value="Listo para procesar")
        self.state_label = tk.Label(self.root, textvariable=self.state_label_var,
                                   font=('Helvetica', 10), fg='blue')
        self.state_label.grid(row=10, column=0, columnspan=3, padx=10, pady=5, sticky='ew')

        # Botón de procesamiento
        self.procesar_btn = tk.Button(self.root, text="Procesar", command=self.iniciar_proceso,
                                     bg='#4CAF50', fg='white', height=2)
        self.procesar_btn.grid(row=11, column=0, columnspan=3, pady=20, sticky='ew', padx=20)

        # Registro de actividad
        tk.Label(self.root, text="Registro de Actividad:", anchor='w').grid(
            row=12, column=0, columnspan=3, sticky='w', padx=10, pady=5)
        self.status_text = tk.Text(self.root, height=12, width=70)
        self.status_text.grid(row=13, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')

        # Agregar scrollbar
        scrollbar = tk.Scrollbar(self.root, command=self.status_text.yview)
        scrollbar.grid(row=13, column=3, sticky='ns')
        self.status_text.config(yscrollcommand=scrollbar.set)

        # Configurar colores para tipos de mensajes
        self.status_text.tag_config("error", foreground="red")
        self.status_text.tag_config("warning", foreground="orange")
        self.status_text.tag_config("success", foreground="green")
        self.status_text.tag_config("time", foreground="blue")

        # Configurar fila para expandir texto de estado
        self.root.rowconfigure(13, weight=1)

    def _generar_opciones_personal(self, lista_personal, campos):
        """Genera opciones formateadas para combobox de personal"""
        opciones = []
        for persona in lista_personal:
            etiqueta = f"{persona[campos[0]]} - {persona[campos[1]]} ({persona[campos[2]]})"
            opciones.append(etiqueta)
        return opciones

    # Métodos de la interfaz
    def select_excel_file(self):
        """Abre el diálogo para seleccionar un archivo Excel"""
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            self.entry_excel_path.delete(0, tk.END)
            self.entry_excel_path.insert(0, file_path)

    def select_date(self, entry_widget):
        """Abre el selector de fecha para un widget de entrada"""
        DateSelector(self.root, entry_widget)

    def update_status(self, message, level="info"):
        """Actualiza el registro de actividad con un nuevo mensaje"""
        if level == "error":
            tag = "error"
            prefix = "❌ ERROR: "
        elif level == "warning":
            tag = "warning"
            prefix = "⚠️ AVISO: "
        elif level == "success":
            tag = "success"
            prefix = "✅ "
        elif level == "time":
            tag = "time"
            prefix = "⏱️ "
        else:
            tag = "info"
            prefix = ""

        self.status_text.insert(tk.END, f"{prefix}{message}\n", tag)
        self.status_text.see(tk.END)  # Auto-scroll al final
        self.root.update_idletasks()

        # También logueamos el mensaje
        if level == "error":
            logger.error(message)
        elif level == "warning":
            logger.warning(message)
        else:
            logger.info(message)

    def set_processing_state(self, is_processing, message="Procesando..."):
        """Actualiza el estado de procesamiento de la interfaz"""
        if is_processing:
            self.state_label_var.set(message)
            self.state_label.config(fg='blue')
            self.procesar_btn.config(state=tk.DISABLED)
        else:
            self.state_label_var.set("Proceso completado")
            self.state_label.config(fg='green')
            self.procesar_btn.config(state=tk.NORMAL)
        self.root.update_idletasks()

    def medir_tiempo(self, operacion, reiniciar=False):
        """Mide y registra el tiempo de operaciones"""
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

            self.update_status(f"Operación '{operacion}' completada en {tiempo_transcurrido:.2f} segundos", "time")

        self.tiempo_inicio = tiempo_actual
        return tiempo_transcurrido

    def iniciar_proceso(self):
        """Método principal que inicia el procesamiento secuencial"""
        # Validar campos
        if not self._validar_campos():
            return

        # Preparar la interfaz
        self.status_text.delete(1.0, tk.END)  # Limpiar log
        self.set_processing_state(True, "Iniciando procesamiento...")

        # Reiniciar estadísticas
        self.facturas_procesadas = 0
        self.facturas_con_error = 0
        self.partidas_procesadas = 0

        # Reiniciar medición de tiempo
        self.medir_tiempo(None, True)

        try:
            # Obtener datos de la interfaz
            datos_comunes = self._obtener_datos_comunes()

            # Procesar el archivo Excel
            self.update_status("Leyendo archivo Excel de partidas...")
            partidas = self.excel_reader.read_partidas(datos_comunes['excel_path'])
            self.medir_tiempo("Lectura de Excel")

            self.update_status(f"Se encontraron {len(partidas)} partidas en el archivo.", "success")

            # Procesar cada partida secuencialmente
            for i, partida in enumerate(partidas, 1):
                self.update_status(f"\n--- Procesando partida {i}/{len(partidas)}: {partida['numero']} ---")
                self.set_processing_state(True, f"Procesando partida {i}/{len(partidas)}...")

                # Verificar directorio de la partida
                partida_dir = os.path.join(datos_comunes['base_dir'], partida['numero'])
                if not os.path.exists(partida_dir):
                    self.update_status(f"Directorio para partida {partida['numero']} no encontrado.", "warning")
                    continue

                # Procesar la partida
                self._procesar_partida(partida, partida_dir, datos_comunes)
                self.partidas_procesadas += 1

                # Actualizar UI después de cada partida
                self.root.update()

            # Proceso completado
            self._mostrar_resumen_final()

        except Exception as e:
            self.update_status(f"Error general en el procesamiento: {str(e)}", "error")
            logger.exception("Error no controlado en el procesamiento")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
        finally:
            # Restaurar interfaz
            self.set_processing_state(False)

    def _validar_campos(self):
        """Valida los campos obligatorios de la interfaz"""
        # Obtener parámetros de la interfaz
        excel_path = self.entry_excel_path.get()
        fecha_documento = self.entry_fecha_documento.get()
        personal_recibio = self.obtener_datos_personal_recibio()
        personal_vobo = self.obtener_datos_personal_vobo()

        # Validar campos requeridos
        if not excel_path:
            messagebox.showerror("Error", "Por favor, seleccione un archivo Excel.")
            return False

        if not fecha_documento:
            messagebox.showerror("Error", "La fecha de elaboración del documento es obligatoria.")
            return False

        if not personal_recibio or not personal_vobo:
            messagebox.showerror(
                "Error",
                "Por favor, seleccione el personal que recibió la compra y el que dio el visto bueno."
            )
            return False

        # Validar formato de fecha
        try:
            datetime.strptime(fecha_documento, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Error", "El formato de fecha debe ser YYYY-MM-DD")
            return False

        return True

    def _obtener_datos_comunes(self):
        """Recopila los datos comunes para el procesamiento"""
        excel_path = self.entry_excel_path.get()
        fecha_documento = self.entry_fecha_documento.get()

        # Convertir fecha a formato de texto
        try:
            fecha_documento_texto = convert_fecha_to_texto(fecha_documento)
        except ValueError as e:
            messagebox.showerror("Error", f"Error en formato de fecha: {str(e)}")
            return None

        datos = {
            'excel_path': excel_path,
            'fecha_documento': fecha_documento,
            'fecha_documento_texto': fecha_documento_texto,
            'mes_asignado': self.mes_asignado_var.get(),
            'personal_recibio': self.obtener_datos_personal_recibio(),
            'personal_vobo': self.obtener_datos_personal_vobo(),
            'base_dir': os.path.dirname(excel_path)
        }

        return datos

    def obtener_datos_personal_recibio(self):
        """Obtiene los datos completos de la persona que recibió seleccionada"""
        if not self.personal_recibio_var.get():
            return None

        etiqueta_seleccionada = self.personal_recibio_var.get()
        for persona in self.lista_personal_recibe:
            etiqueta = (f"{persona['Grado_recibio_la_compra']} - "
                       f"{persona['Nombre_recibio_la_compra']} "
                       f"({persona['Matricula_recibio_la_compra']})")
            if etiqueta == etiqueta_seleccionada:
                return persona
        return None

    def obtener_datos_personal_vobo(self):
        """Obtiene los datos completos de la persona que dio el visto bueno seleccionada"""
        if not self.personal_vobo_var.get():
            return None

        etiqueta_seleccionada = self.personal_vobo_var.get()
        for persona in self.lista_personal_visto_bueno:
            etiqueta = (f"{persona['Grado_Vo_Bo']} - "
                       f"{persona['Nombre_Vo_Bo']} "
                       f"({persona['Matricula_Vo_Bo']})")
            if etiqueta == etiqueta_seleccionada:
                return persona
        return None

    def _procesar_partida(self, partida, partida_dir, datos_comunes):
        """Procesa una partida y todas sus facturas"""
        # Formatear monto de la partida
        monto_formateado = "$ {:,.2f}".format(partida['monto'])

        try:
            # Buscar facturas XML en la partida
            self.medir_tiempo(None)  # Reiniciar contador para esta partida

            # Verificar si hay XML directamente en la carpeta de partida
            xml_files_in_partida = [
                f for f in os.listdir(partida_dir)
                if f.lower().endswith('.xml') and os.path.isfile(os.path.join(partida_dir, f))
            ]

            facturas_procesadas = 0
            facturas_con_error = 0
            facturas_info = []

            if xml_files_in_partida:
                # CASO 1: XML directamente en la carpeta de partida (una sola factura)
                self.update_status(f"📄 Encontrado XML directamente en la carpeta de partida")

                # Procesar la factura única
                xml_file = os.path.join(partida_dir, xml_files_in_partida[0])
                resultado = self._procesar_factura(xml_file, partida_dir, partida, monto_formateado, datos_comunes)

                if resultado:
                    facturas_procesadas += 1
                    facturas_info.append(resultado)
                else:
                    facturas_con_error += 1
            else:
                # CASO 2: Buscar en subcarpetas (múltiples facturas)
                subdirs = [
                    d for d in os.listdir(partida_dir)
                    if os.path.isdir(os.path.join(partida_dir, d))
                ]

                self.update_status(f"📂 Partida {partida['numero']}: {len(subdirs)} subcarpetas encontradas.")

                # Procesar cada subcarpeta que contenga XML
                for subdir in subdirs:
                    factura_dir = os.path.join(partida_dir, subdir)
                    xml_files = [
                        f for f in os.listdir(factura_dir)
                        if f.lower().endswith('.xml') and os.path.isfile(os.path.join(factura_dir, f))
                    ]

                    if xml_files:
                        self.update_status(f"  - Procesando factura en {subdir}...")

                        # Procesar la factura
                        xml_file = os.path.join(factura_dir, xml_files[0])
                        resultado = self._procesar_factura(xml_file, factura_dir, partida,
                        monto_formateado, datos_comunes)

                        if resultado:
                            facturas_procesadas += 1
                            facturas_info.append(resultado)
                        else:
                            facturas_con_error += 1

                    # Actualizar la UI después de cada factura
                    self.root.update()

            # Actualizar estadísticas globales
            self.facturas_procesadas += facturas_procesadas
            self.facturas_con_error += facturas_con_error

            # Generar relación de facturas si hay información disponible
            if facturas_info:
                self._generar_relacion_facturas(partida, facturas_info, partida_dir, datos_comunes)

            # Resumen de la partida
            self.medir_tiempo(f"Partida {partida['numero']} completa")
            self.update_status(
                f"Partida {partida['numero']} completada: {facturas_procesadas} facturas procesadas, "
                f"{facturas_con_error} con errores.",
                "success" if facturas_con_error == 0 else "warning"
            )

            return {
                'numero': partida['numero'],
                'descripcion': partida['descripcion'],
                'facturas_procesadas': facturas_procesadas,
                'facturas_con_error': facturas_con_error
            }

        except Exception as e:
            self.update_status(f"Error al procesar partida {partida['numero']}: {str(e)}", "error")
            logger.exception(f"Error procesando partida {partida['numero']}")
            return None

    def _procesar_factura(self, xml_file, output_dir, partida, monto_formateado, datos_comunes):
        """Procesa una factura individual"""
        try:
            self.update_status(f"🔍 Analizando XML: {os.path.basename(xml_file)}...")

            # Medir tiempo de procesamiento XML
            self.medir_tiempo(None)

            # 1. Extraer información base del XML
            xml_data = self.xml_processor.read_xml(xml_file)
            tiempo_xml = self.medir_tiempo("Lectura XML")

            if not xml_data:
                self.update_status(f"Error: No se pudo extraer información del XML", "error")
                return None

            # 2. Crear el diccionario data completo
            data = self._crear_diccionario_datos_completo(
                xml_data,
                partida,
                monto_formateado,
                datos_comunes
            )

            # 3. Pre-procesar conceptos (formatearlos automáticamente)
            conceptos_str = self._formatear_conceptos(data['Conceptos'])

            # 4. Si está habilitado el editor de conceptos, mostrarlo
            if self.config.get('usar_editor_conceptos', True):
                self.update_status(f"✏️ Abriendo editor de conceptos...")

                # Este es un punto crítico donde debemos esperar la interacción del usuario
                conceptos_editados = self._editar_conceptos_bloqueante(data['Conceptos'], partida['descripcion'])

                if conceptos_editados:
                    data['Empleo_recurso'] = conceptos_editados
                else:
                    # Si no se editaron, usar los conceptos formateados automáticamente
                    data['Empleo_recurso'] = conceptos_str
            else:
                # Usar el formato automático
                data['Empleo_recurso'] = conceptos_str

            # 5. Generar documentos
            self.update_status(f"📝 Generando documentos...")
            self.medir_tiempo(None)
            documentos_generados = self.document_generator.generate_all_documents(data, output_dir)
            self.medir_tiempo("Generación documentos")

            # 6. Registro de éxito
            self.update_status(
                f"✅ Documentos generados para factura {data['Serie']}{data['Numero']}",
                "success"
            )

            # 7. Retornar información para registro y relación
            return {
                'serie_numero': f"{data['Serie']}{data['Numero']}",
                'fecha': data.get('Fecha_factura_texto', data.get('Fecha_factura', '')),
                'emisor': data['Nombre_Emisor'],
                'rfc_emisor': data['Rfc_emisor'],
                'monto': data['monto'],
                'conceptos': data.get('Empleo_recurso', ''),
                'documentos': documentos_generados
            }

        except Exception as e:
            self.update_status(
                f"Error al procesar factura {os.path.basename(xml_file)}: {str(e)}",
                "error"
            )
            logger.exception(f"Error procesando factura {xml_file}")
            return None

    def _editar_conceptos_bloqueante(self, conceptos, descripcion_partida):
        """
        Abre el editor de conceptos y espera de forma bloqueante hasta que el usuario confirme

        Args:
            conceptos: Diccionario de conceptos {descripcion: cantidad}
            descripcion_partida: Descripción de la partida

        Returns:
            str: Conceptos editados o None si se canceló
        """
        # Usar la función simplificada que ya maneja todo internamente
        return editar_conceptos_simple(self.root, conceptos, descripcion_partida)

    def _crear_diccionario_datos_completo(self, xml_data, partida, monto_formateado, datos_comunes):
        """
        Crea un diccionario completo combinando todas las fuentes de datos
        """
        # Crear un nuevo diccionario
        data = {}

        # 1. Agregar datos del XML
        data.update(xml_data)

        # 2. Agregar datos de la interfaz principal
        data['Fecha_doc'] = datos_comunes['fecha_documento_texto']
        data['Mes'] = datos_comunes['mes_asignado']

        # 3. Agregar información de la partida
        data['No_partida'] = partida['numero']
        data['Descripcion_partida'] = partida['descripcion']
        data['monto'] = monto_formateado

        # 4. Agregar información del personal seleccionado
        for key, value in datos_comunes['personal_recibio'].items():
            data[key] = value

        for key, value in datos_comunes['personal_vobo'].items():
            data[key] = value

        # 5. Información de fechas formateadas
        # Convertir ISO fecha a formato legible
        if 'Fecha_ISO' in data:
            fecha_obj = datetime.strptime(data['Fecha_ISO'].split('T')[0], '%Y-%m-%d')
            data['Fecha_original'] = data['Fecha_ISO']

            # Formato numérico para algunas ocasiones donde se necesite
            data['Fecha_factura'] = fecha_obj.strftime('%d/%m/%Y')

            # Formato textual en español
            try:
                locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Intentar configurar locale español
            except:
                try:
                    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # Alternativa para Windows
                except:
                    pass  # Si no se puede establecer, usar la configuración por defecto

            # Usar format_date de babel para formatear con mes en texto
            data['Fecha_factura_texto'] = format_date(fecha_obj, format="d 'de' MMMM 'del' yyyy", locale='es')
            # Capitalizar primera letra
            if data['Fecha_factura_texto']:
                data['Fecha_factura_texto'] = data['Fecha_factura_texto'][0].upper() + data['Fecha_factura_texto'][1:]

        # 6. Información del Folio Fiscal
        if 'UUid' in data:
            data['Folio_Fiscal'] = data['UUid']

        # 7. Generar número de oficio
        data['No_of_remision'] = partida['numero_adicional']

        # 8. Auto-generar No_mensaje y Fecha_mensaje
        data['Fecha_mensaje'] = format_fecha_mensaje(datos_comunes['fecha_documento'])

        return data

    def _formatear_conceptos(self, conceptos):
        """
        Formatea los conceptos para presentación
        """
        if not conceptos:
            return "Conceptos no disponibles"

        return f"{conceptos}"

    # aqui inicia

    def _generar_relacion_facturas(self, partida, facturas_info, partida_dir, datos_nivel1):
        """
    Genera un documento de relación de facturas para la partida

    Args:
        partida: Datos de la partida
        facturas_info: Lista de información de facturas procesadas
        partida_dir: Directorio de la partida
        datos_nivel1: Datos del nivel 1
    """
        try:
            self.update_status(f"Generando relación de facturas para partida {partida['numero']}...")

            # Importar el módulo de plantillas de partidas
            from generators.plantillas_partidas import procesar_plantillas_partida, calcular_montos_facturas

            # Preparar datos comunes para las plantillas
            datos_comunes = {
                'mes_asignado': datos_nivel1['mes_asignado'],
                'fecha_documento': datos_nivel1['fecha_documento'],
                'fecha_documento_texto': datos_nivel1['fecha_documento_texto'],
                'personal_recibio': datos_nivel1['personal_recibio'],
                'personal_vobo': datos_nivel1['personal_vobo']
            }

            # Calcular información resumida de facturas (totales, montos, etc.)
            info_facturas = calcular_montos_facturas(facturas_info)
            self.update_status(f"  - Calculados totales para {info_facturas['total_facturas']} facturas. "
                            f"Monto total: {info_facturas['monto_formateado']}")

            # Añadir la información resumida a los datos comunes
            datos_comunes['info_facturas'] = info_facturas

            # Procesar todas las plantillas de la partida
            self.update_status("Procesando plantillas de documentos...")
            archivos_generados = procesar_plantillas_partida(
                partida,
                facturas_info,
                partida_dir,
                datos_comunes
            )

            # Registrar los archivos generados
            if archivos_generados:
                for tipo, ruta in archivos_generados.items():
                    self.update_status(f"  - {tipo.capitalize()}: {os.path.basename(ruta)}")

            self.update_status(
                f"✅ Relación de facturas generada para partida {partida['numero']}",
                "success"
            )

            return archivos_generados
        except Exception as e:
            self.update_status(
                f"Error al generar relación de facturas: {str(e)}",
                "error"
            )
            import traceback
            traceback.print_exc()
            return None


# aqui termina el codigo

    def _mostrar_resumen_final(self):
        """Muestra el resumen final del procesamiento"""
        tiempo_total = sum(self.tiempos_operaciones.values())

        self.update_status("\n===== RESUMEN DEL PROCESAMIENTO =====")
        self.update_status(f"Total partidas procesadas: {self.partidas_procesadas}")
        self.update_status(f"Total facturas procesadas: {self.facturas_procesadas}")
        self.update_status(f"Facturas con error: {self.facturas_con_error}")
        self.update_status(f"Tiempo total de procesamiento: {tiempo_total:.2f} segundos")

        # Mostrar tiempos por tipo de operación
        if self.tiempos_operaciones:
            self.update_status("\nTiempos por operación:")
            for operacion, tiempo in sorted(self.tiempos_operaciones.items(), key=lambda x: x[1], reverse=True):
                if "completa" not in operacion:  # Excluir totales de partidas que ya están sumados en otros
                    porcentaje = (tiempo / tiempo_total) * 100
                    self.update_status(f"  - {operacion}: {tiempo:.2f} segundos ({porcentaje:.1f}%)", "time")

        # Mensaje final
        mensaje_final = f"Proceso completado. {self.facturas_procesadas} facturas procesadas en {self.partidas_procesadas} partidas."
        self.update_status(mensaje_final, "success")
        messagebox.showinfo("Proceso Completado", mensaje_final)

# ===================================================
# FUNCIÓN PRINCIPAL
# ===================================================

def main():
    """Función principal de la aplicación"""
    root = tk.Tk()
    app = AutomatizacionApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
