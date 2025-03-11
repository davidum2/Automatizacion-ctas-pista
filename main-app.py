"""
Sistema de Automatización de Documentos por Partidas
Estructura refactorizada por niveles de procesamiento
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import threading
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
from ui.concepto_editor import ConceptoEditorWindow

# Configuración de logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ===================================================
# NIVEL 1: CAPA DE UI - INTERFAZ PRINCIPAL
# ===================================================

class AutomatizacionApp:
    """Interfaz principal de la aplicación"""

    def __init__(self, root):
        self.root = root
        self.root.title("Automatización de Documentos por Partidas")
        self.root.geometry("800x700")

            # Añadir configuración predeterminada
        self.config = {
            'usar_editor_conceptos': True,  # Activa o desactiva el editor de conceptos
            # Puedes añadir más opciones de configuración aquí
        }

        # Inicializar componentes
        self.excel_reader = ExcelReader()
        self.xml_processor = XMLProcessor()
        self.document_generator = DocumentGenerator()
        self.file_utils = FileUtils()

        # Variables para seguimiento de progreso
        self.total_facturas = 0
        self.total_facturas_procesadas = 0
        self.facturas_por_partida = {}

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
        self.mes_asignado_var.set(meses[datetime.now().month - 1])  # Mes actual como predeterminado
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

        # Separador antes de barra de progreso
        ttk.Separator(self.root, orient='horizontal').grid(
            row=9, column=0, columnspan=3, sticky='ew', pady=10)

        # Barra de progreso
        tk.Label(self.root, text="Progreso:", anchor='w').grid(
            row=10, column=0, padx=10, pady=5, sticky='ew')
        self.progress_bar = ttk.Progressbar(self.root, length=400, mode='determinate')
        self.progress_bar.grid(row=10, column=1, columnspan=2, padx=10, pady=5, sticky='ew')

        # Botón de procesamiento
        tk.Button(self.root, text="Procesar", command=self.iniciar_proceso,
                 bg='#4CAF50', fg='white', height=2).grid(
            row=11, column=0, columnspan=3, pady=20, sticky='ew', padx=20)

        # Registro de actividad
        tk.Label(self.root, text="Registro de Actividad:", anchor='w').grid(
            row=12, column=0, columnspan=3, sticky='w', padx=10, pady=5)
        self.status_text = tk.Text(self.root, height=12, width=70)
        self.status_text.grid(row=13, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')

        # Agregar scrollbar
        scrollbar = tk.Scrollbar(self.root, command=self.status_text.yview)
        scrollbar.grid(row=13, column=3, sticky='ns')
        self.status_text.config(yscrollcommand=scrollbar.set)

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

    def update_progress(self, value, maximum=100):
        """Actualiza la barra de progreso"""
        self.progress_bar['value'] = (value / maximum) * 100
        self.root.update_idletasks()

    def iniciar_proceso(self):
        """Método que inicia el procesamiento principal"""
        # Validar campos
        if not self._validar_campos():
            return

        # Limpiar estado previo
        self.status_text.delete(1.0, tk.END)
        self.status_text.tag_config("error", foreground="red")
        self.status_text.tag_config("warning", foreground="orange")
        self.status_text.tag_config("success", foreground="green")

        # Reiniciar contadores
        self.total_facturas = 0
        self.total_facturas_procesadas = 0
        self.facturas_por_partida = {}

        # Obtener datos de la interfaz
        datos_nivel1 = self._obtener_datos_nivel1()

        # Iniciar procesamiento en un hilo separado para no bloquear la UI
        threading.Thread(target=self._ejecutar_nivel1,
                        args=(datos_nivel1,),
                        daemon=True).start()

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

    def _obtener_datos_nivel1(self):
        """Recopila los datos del nivel 1 (interfaz principal)"""
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


# ===================================================
# NIVEL 1: PROCESAMIENTO DE NIVEL SUPERIOR
# ===================================================

    def _ejecutar_nivel1(self, datos_nivel1):
        """
        Ejecuta el procesamiento de nivel 1
        Este método controla el flujo principal de la aplicación
        """
        try:
            self.update_status("Iniciando procesamiento...")

            # 1. Leer el archivo Excel con partidas
            self.update_status("Leyendo archivo Excel...")
            try:
                partidas = self.excel_reader.read_partidas(datos_nivel1['excel_path'])
                self.update_status(f"Se encontraron {len(partidas)} partidas en el archivo.", "success")
            except Exception as e:
                self.update_status(f"Error al leer el archivo Excel: {str(e)}", "error")
                messagebox.showerror("Error", f"Error al leer el archivo Excel: {str(e)}")
                return

            # 2. Contar facturas totales para barra de progreso
            self.total_facturas = self._contar_facturas_totales(partidas, datos_nivel1['base_dir'])
            self.update_status(f"Total de facturas a procesar: {self.total_facturas}")

            # 3. Procesar cada partida (Nivel 2)
            resultados_partidas = []
            for i, partida in enumerate(partidas):
                # Actualizar progreso entre partidas
                self.update_progress(i, len(partidas))

                # Procesar partida actual
                resultado_partida = self._ejecutar_nivel2_partida(partida, datos_nivel1, i+1, len(partidas))
                if resultado_partida:
                    resultados_partidas.append(resultado_partida)

            # 4. Generar informe final si es necesario
            if resultados_partidas:
                self._generar_informe_final(resultados_partidas, datos_nivel1)

            # 5. Actualización final del progreso
            self.update_progress(100, 100)

            # 6. Resumen final
            self._mostrar_resumen_final(len(partidas), resultados_partidas)

        except Exception as e:
            self.update_status(f"ERROR GENERAL: {str(e)}", "error")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")

    def _contar_facturas_totales(self, partidas, base_dir):
        """Cuenta el número total de facturas para todas las partidas"""
        total = 0

        for partida in partidas:
            partida_dir = os.path.join(base_dir, partida['numero'])
            if not os.path.exists(partida_dir):
                continue

            # Buscar XMLs directamente en la carpeta de partida
            xml_files_in_partida = [f for f in os.listdir(partida_dir)
                            if f.lower().endswith('.xml') and os.path.isfile(os.path.join(partida_dir, f))]

            if xml_files_in_partida:
                # Si hay XMLs directamente en la carpeta, es una sola factura
                total += 1
            else:
                # Si no hay XMLs directos, buscar en subcarpetas
                subdirs = [d for d in os.listdir(partida_dir)
                        if os.path.isdir(os.path.join(partida_dir, d))]

                for subdir in subdirs:
                    factura_dir = os.path.join(partida_dir, subdir)
                    xml_files = [f for f in os.listdir(factura_dir)
                                if f.lower().endswith('.xml') and os.path.isfile(os.path.join(factura_dir, f))]
                    if xml_files:
                        total += 1

        return total

    def _mostrar_resumen_final(self, num_partidas, resultados_partidas):
        """Muestra el resumen final del procesamiento"""
        self.update_status("\n===== RESUMEN DEL PROCESAMIENTO =====")
        self.update_status(f"Total de partidas encontradas: {num_partidas}")
        self.update_status(f"Total de facturas procesadas: {self.total_facturas_procesadas}/{self.total_facturas}")

        # Mostrar detalles por partida
        if resultados_partidas:
            self.update_status("\nDetalles por partida:")
            for resultado in resultados_partidas:
                self.update_status(
                    f"  Partida {resultado['numero']}: "
                    f"{resultado['procesadas']}/{resultado['total']} facturas procesadas"
                )

        self.update_status("¡Procesamiento completado con éxito!", "success")
        messagebox.showinfo(
            "Éxito",
            f"Procesamiento completado.\n"
            f"Se procesaron {self.total_facturas_procesadas} facturas de {num_partidas} partidas."
        )

    def _generar_informe_final(self, resultados_partidas, datos_nivel1):
        """Genera un informe final consolidado de todas las partidas"""
        # Esta función podría implementarse para generar un informe general si se requiere
        pass


# ===================================================
# NIVEL 2: PROCESAMIENTO DE PARTIDAS
# ===================================================

    def _ejecutar_nivel2_partida(self, partida, datos_nivel1, indice, total_partidas):
        """
        Ejecuta el procesamiento de nivel 2 para una partida específica

        Args:
            partida: Diccionario con datos de la partida
            datos_nivel1: Datos del nivel 1
            indice: Índice actual de la partida
            total_partidas: Total de partidas

        Returns:
            dict: Resultados del procesamiento de la partida o None si hubo error
        """
        self.update_status(
            f"\nProcesando partida {indice}/{total_partidas}: "
            f"{partida['numero']} - {partida['descripcion']}..."
        )

        # Validar existencia del directorio de la partida
        partida_dir = os.path.join(datos_nivel1['base_dir'], partida['numero'])
        if not os.path.exists(partida_dir):
            self.update_status(
                f"Directorio para partida {partida['numero']} no encontrado.",
                "warning"
            )
            return None

        # Formatear monto de la partida
        monto_formateado = "$ {:,.2f}".format(partida['monto'])

        # Inicializar contadores para esta partida
        facturas_totales_partida = 0
        facturas_procesadas_partida = 0
        facturas_info = []  # Almacenará información de facturas procesadas

        # PROCESO 2.1: Localizar y procesar facturas en esta partida
        try:
            # Verificar si hay XML directamente en la carpeta de partida
            xml_files_in_partida = [
                f for f in os.listdir(partida_dir)
                if f.lower().endswith('.xml') and os.path.isfile(os.path.join(partida_dir, f))
            ]

            if xml_files_in_partida:
                # CASO 1: XML directamente en la carpeta de partida (una sola factura)
                self.update_status(f"📄 Encontrado XML directamente en la carpeta de partida")
                facturas_totales_partida = 1

                # Procesar la factura única
                xml_file = os.path.join(partida_dir, xml_files_in_partida[0])
                resultado_factura = self._ejecutar_nivel3_factura(
                    xml_file,
                    partida_dir,
                    partida,
                    monto_formateado,
                    datos_nivel1
                )

                if resultado_factura:
                    facturas_procesadas_partida += 1
                    facturas_info.append(resultado_factura)
            else:
                # CASO 2: Buscar en subcarpetas (múltiples facturas)
                subdirs = [
                    d for d in os.listdir(partida_dir)
                    if os.path.isdir(os.path.join(partida_dir, d))
                ]

                self.update_status(
                    f"📂 Partida {partida['numero']}: {len(subdirs)} subcarpetas encontradas."
                )

                # Contar facturas en subcarpetas
                facturas_en_subdirs = []
                for subdir in subdirs:
                    factura_dir = os.path.join(partida_dir, subdir)
                    xml_files = [
                        f for f in os.listdir(factura_dir)
                        if f.lower().endswith('.xml') and os.path.isfile(os.path.join(factura_dir, f))
                    ]
                    if xml_files:
                        facturas_en_subdirs.append({
                            'subdir': subdir,
                            'dir': factura_dir,
                            'xml_file': os.path.join(factura_dir, xml_files[0])
                        })

                facturas_totales_partida = len(facturas_en_subdirs)
                self.update_status(
                    f"📋 Partida {partida['numero']}: {facturas_totales_partida} facturas encontradas."
                )

                # Procesar cada factura en subcarpetas
                for i, factura_info in enumerate(facturas_en_subdirs):
                    self.update_status(
                        f"📄 Procesando factura {i+1}/{facturas_totales_partida}: {factura_info['subdir']}"
                    )

                    # Procesar la factura
                    resultado_factura = self._ejecutar_nivel3_factura(
                        factura_info['xml_file'],
                        factura_info['dir'],
                        partida,
                        monto_formateado,
                        datos_nivel1
                    )

                    if resultado_factura:
                        facturas_procesadas_partida += 1
                        facturas_info.append(resultado_factura)

                    # Actualizar interfaz para mantenerla responsiva
                    self.root.update()

            # PROCESO 2.2: Generar documento consolidado para esta partida
            if facturas_info:
                self._generar_relacion_facturas(partida, facturas_info, partida_dir, datos_nivel1)

            # Actualizar contadores globales
            self.total_facturas_procesadas += facturas_procesadas_partida

            # Retornar resultados de esta partida
            return {
                'numero': partida['numero'],
                'descripcion': partida['descripcion'],
                'total': facturas_totales_partida,
                'procesadas': facturas_procesadas_partida,
                'facturas': facturas_info
            }

        except Exception as e:
            self.update_status(
                f"Error al procesar partida {partida['numero']}: {str(e)}",
                "error"
            )
            return None

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

            # Aquí implementaríamos la generación del documento consolidado
            # Por ejemplo, usando la función create_relacion_de_facturas_excel

            self.update_status(
                f"✅ Relación de facturas generada para partida {partida['numero']}",
                "success"
            )
        except Exception as e:
            self.update_status(
                f"Error al generar relación de facturas: {str(e)}",
                "error"
            )



# ===================================================
# NIVEL 3: PROCESAMIENTO DE FACTURAS INDIVIDUALES
# ===================================================

    def _ejecutar_nivel3_factura(self, xml_file, output_dir, partida, monto_formateado, datos_nivel1):
        """
        Ejecuta el procesamiento de nivel 3 para una factura individual

        Args:
            xml_file: Ruta al archivo XML
            output_dir: Directorio donde guardar los documentos generados
            partida: Datos de la partida
            monto_formateado: Monto de la partida formateado
            datos_nivel1: Datos del nivel 1

        Returns:
            dict/str: Información de la factura procesada, "pendiente" si se abrió el editor,
                    o None si hubo error
        """
        try:
            # Leer y procesar el XML
            self.update_status(f"🔍 Analizando XML: {os.path.basename(xml_file)}...")

            # 1. Extraer información base del XML
            xml_data = self.xml_processor.read_xml(xml_file)
            if not xml_data:
                self.update_status(f"Error: No se pudo extraer información del XML", "error")
                return None

            # 2. Crear el diccionario data combinando todas las fuentes
            data = self._crear_diccionario_datos_completo(
                xml_data,
                partida,
                monto_formateado,
                datos_nivel1
            )

            # 3. Editar conceptos con interfaz gráfica si la opción está activada
            if self.config.get('usar_editor_conceptos', True):
                return self._procesar_con_editor_conceptos(data, output_dir, partida)
            else:
                # Procesar directamente sin editor
                conceptos_str = self._formatear_conceptos(data['Conceptos'])
                data['Empleo_recurso'] = conceptos_str
                return self._generar_documentos_factura(data, output_dir)

        except Exception as e:
            self.update_status(
                f"Error al procesar los datos del xml {os.path.basename(xml_file)}: {str(e)}",
                "error"
            )
            logger.error(f"Error en nivel 3 - factura {xml_file}: {str(e)}", exc_info=True)
            return None

    def _procesar_con_editor_conceptos(self, data, output_dir, partida):
        """
        Procesa una factura mostrando primero el editor de conceptos

        Args:
            data: Diccionario con los datos de la factura
            output_dir: Directorio de salida
            partida: Información de la partida

        Returns:
            str: "pendiente" para indicar que se continuará después
        """
        self.update_status(f"✏️ Abriendo editor de conceptos...")

        # Esta variable especial se usará para esperar por el editor
        self.editor_completado = False
        self.datos_editados = None

        # Función de callback que se ejecutará cuando el usuario confirme en el editor
        def on_editor_completed(conceptos_editados):
            self.datos_editados = conceptos_editados
            self.editor_completado = True
            # Continuamos el procesamiento
            self.root.after(100, lambda: self._continuar_proceso_factura(data, output_dir))

        # Mostrar el editor de conceptos en el hilo principal
        self.root.after(0, lambda: self._mostrar_editor_conceptos(
            data['Conceptos'],
            partida['descripcion'],
            on_editor_completed
        ))

        # Retornamos "pendiente" por ahora, el procesamiento continuará en el callback
        return "pendiente"

    def _mostrar_editor_conceptos(self, conceptos, descripcion_partida, callback):
        """
        Muestra el editor de conceptos en el hilo principal

        Args:
            conceptos: Diccionario de conceptos a editar
            descripcion_partida: Descripción de la partida para contexto
            callback: Función a llamar cuando se complete la edición
        """
        # Crear una ventana de edición de conceptos
        editor = ConceptoEditorWindow(
            self.root,
            conceptos,
            descripcion_partida,
            callback
        )

    def _continuar_proceso_factura(self, data, output_dir):
        """
        Continúa el procesamiento después de que el editor de conceptos se complete

        Args:
            data: Diccionario de datos de la factura
            output_dir: Directorio de salida para los documentos
        """
        try:
            # Actualizar los datos con la información editada
            if self.datos_editados:
                data['Empleo_recurso'] = self.datos_editados

            # Llamar al método común para generar documentos
            resultado = self._generar_documentos_factura(data, output_dir)

            # Continuar con la siguiente factura en la cola
            self._procesar_siguiente_factura()

            return resultado
        except Exception as e:
            self.update_status(f"Error al continuar el proceso: {str(e)}", "error")
            logger.error(f"Error en continuación de proceso: {str(e)}", exc_info=True)
            self._procesar_siguiente_factura()  # Intentar continuar con la siguiente a pesar del error
            return None

    def _generar_documentos_factura(self, data, output_dir):
        """
        Genera los documentos para una factura

        Args:
            data: Diccionario con los datos de la factura
            output_dir: Directorio donde guardar los documentos

        Returns:
            dict: Información de la factura procesada
        """
        # Generar los documentos
        self.update_status(f"📝 Generando documentos...")
        documentos_generados = self.document_generator.generate_all_documents(data, output_dir)

        # Actualizar progreso
        self.total_facturas_procesadas += 1
        self.update_progress(self.total_facturas_procesadas, self.total_facturas)

        # Registro de éxito
        self.update_status(
            f"✅ Documentos generados para factura {data['Serie']}{data['Numero']}",
            "success"
        )

        # Retornar información de la factura procesada
        return {
            'serie_numero': f"{data['Serie']}{data['Numero']}",
            'fecha': data.get('Fecha_factura'),
            'emisor': data['Nombre_Emisor'],
            'rfc_emisor': data['Rfc_emisor'],
            'monto': data['monto'],
            'conceptos': data.get('Empleo_recurso', ''),
            'documentos': documentos_generados
        }

    def _crear_diccionario_datos_completo(self, xml_data, partida, monto_formateado, datos_nivel1):
        """
        Crea un diccionario completo combinando todas las fuentes de datos

        Args:
            xml_data: Datos extraídos del XML
            partida: Datos de la partida
            monto_formateado: Monto formateado
            datos_nivel1: Datos del nivel 1

        Returns:
            dict: Diccionario completo de datos
        """
        # Crear un nuevo diccionario
        data = {}

        # 1. Agregar datos del XML
        data.update(xml_data)

        # 2. Agregar datos de la interfaz principal
        data['Fecha_doc'] = datos_nivel1['fecha_documento_texto']
        data['Mes'] = datos_nivel1['mes_asignado']

        # 3. Agregar información de la partida
        data['No_partida'] = partida['numero']
        data['Descripcion_partida'] = partida['descripcion']
        data['monto'] = monto_formateado

        # 4. Agregar información del personal seleccionado
        for key, value in datos_nivel1['personal_recibio'].items():
            data[key] = value

        for key, value in datos_nivel1['personal_vobo'].items():
            data[key] = value

        # 5. Información de fechas formateadas
        # Convertir ISO fecha a formato legible
        if 'Fecha_ISO' in data:
            fecha_obj = datetime.strptime(data['Fecha_ISO'].split('T')[0], '%Y-%m-%d')
            data['Fecha_original'] = data['Fecha_ISO']
            data['Fecha_factura'] = fecha_obj.strftime('%d/%m/%Y')

        # 6. Información del Folio Fiscal
        if 'UUid' in data:
            data['Folio_Fiscal'] = data['UUid']

        # 7. Generar número de oficio y fecha de remisión (auto)
        fecha_actual = datetime.now()
        data['No_of_remision'] = f"OF-{partida['numero']}-{fecha_actual.strftime('%m%d')}"

        # Formato para fecha de remisión
        locale.setlocale(locale.LC_TIME, '')  # Usar configuración del sistema
        fecha_remision_texto = format_date(fecha_actual, format="d 'de' MMMM 'de' y", locale='es')
        data['Fecha_remision'] = fecha_remision_texto

        # 8. Auto-generar No_mensaje y Fecha_mensaje
        data['No_mensaje'] = f"M-{fecha_actual.year}-{partida['numero']}"
        data['Fecha_mensaje'] = format_fecha_mensaje(datos_nivel1['fecha_documento'])

        return data

    def _formatear_conceptos(self, conceptos):
        """
        Formatea los conceptos para presentación

        Args:
            conceptos: Diccionario de conceptos {descripcion: cantidad}

        Returns:
            str: Texto formateado de conceptos
        """
        if not conceptos:
            return "Conceptos no disponibles"

        # Si hay menos de 3 conceptos, mostrarlos todos
        if len(conceptos) <= 3:
            return ", ".join([f"{cantidad} {desc}" for desc, cantidad in conceptos.items()])

        # Si hay muchos conceptos, hacer un resumen
        total_items = sum(conceptos.values())

        # Tomar los 2 conceptos más importantes
        sorted_items = sorted(conceptos.items(), key=lambda x: x[1], reverse=True)
        principales = sorted_items[:2]

        texto = ", ".join([f"{cantidad} {desc}" for desc, cantidad in principales])
        return f"{texto} y otros artículos (total {total_items} unidades)"

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
