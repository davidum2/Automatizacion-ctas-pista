"""
Módulo para la ventana principal de la aplicación
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import logging
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import APP_CONFIG, PERSONAL_RECIBE, PERSONAL_VISTO_BUENO, MESES
from ui.dialogs import DateSelector
from controllers.process_controller import ProcessController

logger = logging.getLogger(__name__)

class AutomatizacionAppWindow:
    """Clase para la ventana principal de la aplicación"""

    def __init__(self, root):
        """Inicializa la ventana principal"""
        self.root = root
        self.root.title("Automatización de Documentos por Partidas")
        self.root.geometry("800x700")
        
        # Controlador de proceso
        self.process_controller = ProcessController(self)
        
        # Crear la interfaz
        self.create_widgets()
        
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
                  command=lambda: DateSelector(self.root, self.entry_fecha_documento)).grid(
            row=3, column=2, padx=10, pady=5)

        # Mes asignado
        tk.Label(self.root, text="Mes asignado:", anchor='w').grid(
            row=4, column=0, padx=10, pady=5, sticky='ew')
        self.mes_asignado_var = tk.StringVar(self.root)
        self.mes_asignado_var.set(MESES[datetime.now().month - 1])  # Mes actual predeterminado
        option_menu_meses = tk.OptionMenu(self.root, self.mes_asignado_var, *MESES)
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
        opciones_recibio = self._generar_opciones_personal(PERSONAL_RECIBE,
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
        opciones_vobo = self._generar_opciones_personal(PERSONAL_VISTO_BUENO,
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
        self.procesar_btn = tk.Button(self.root, text="Procesar", 
                                     command=self.iniciar_proceso,
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

    def select_excel_file(self):
        """Abre el diálogo para seleccionar un archivo Excel"""
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            self.entry_excel_path.delete(0, tk.END)
            self.entry_excel_path.insert(0, file_path)

    def iniciar_proceso(self):
        """Inicia el proceso de procesamiento"""
        # Recopilar datos de la interfaz
        datos_interfaz = self.recopilar_datos_interfaz()
        
        # Validar datos
        if not datos_interfaz:
            return
            
        # Limpiar la interfaz para un nuevo procesamiento
        self.status_text.delete(1.0, tk.END)
        
        # Configurar la interfaz para procesamiento
        self.set_processing_state(True, "Iniciando procesamiento...")
        
        # Iniciar el procesamiento
        self.process_controller.iniciar_procesamiento(datos_interfaz)

    def recopilar_datos_interfaz(self):
        """Recopila los datos de la interfaz para el procesamiento"""
        # Obtener parámetros de la interfaz
        excel_path = self.entry_excel_path.get()
        fecha_documento = self.entry_fecha_documento.get()
        mes_asignado = self.mes_asignado_var.get()
        
        # Obtener datos de personal
        personal_recibio = self.obtener_datos_personal_recibio()
        personal_vobo = self.obtener_datos_personal_vobo()
        
        # Validar campos requeridos
        if not excel_path:
            messagebox.showerror("Error", "Por favor, seleccione un archivo Excel.")
            return None

        if not fecha_documento:
            messagebox.showerror("Error", "La fecha de elaboración del documento es obligatoria.")
            return None

        if not personal_recibio or not personal_vobo:
            messagebox.showerror(
                "Error",
                "Por favor, seleccione el personal que recibió la compra y el que dio el visto bueno."
            )
            return None

        # Validar formato de fecha
        try:
            datetime.strptime(fecha_documento, APP_CONFIG['formato_fecha'])
        except ValueError:
            messagebox.showerror("Error", f"El formato de fecha debe ser {APP_CONFIG['formato_fecha']}")
            return None
            
        # Estructura de datos para el procesamiento
        return {
            'excel_path': excel_path,
            'fecha_documento': fecha_documento,
            'mes_asignado': mes_asignado, 
            'personal_recibio': personal_recibio,
            'personal_vobo': personal_vobo,
            'base_dir': os.path.dirname(excel_path)
        }

    def obtener_datos_personal_recibio(self):
        """Obtiene los datos completos de la persona que recibió seleccionada"""
        if not self.personal_recibio_var.get():
            return None

        etiqueta_seleccionada = self.personal_recibio_var.get()
        for persona in PERSONAL_RECIBE:
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
        for persona in PERSONAL_VISTO_BUENO:
            etiqueta = (f"{persona['Grado_Vo_Bo']} - "
                       f"{persona['Nombre_Vo_Bo']} "
                       f"({persona['Matricula_Vo_Bo']})")
            if etiqueta == etiqueta_seleccionada:
                return persona
        return None
    
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