import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import threading

# Importar módulos refactorizados
from core.excel_reader import ExcelReader
from core.xml_processor import XMLProcessor
from core.document_generator import DocumentGenerator
from utils.file_utils import FileUtils
from ui.date_selector import DateSelector
from utils.formatters import convert_fecha_to_texto

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

# Importar módulos refactorizados
from core.excel_reader import ExcelReader
from core.xml_processor import XMLProcessor
from core.document_generator import DocumentGenerator
from utils.file_utils import FileUtils
from ui.date_selector import DateSelector
from utils.formatters import convert_fecha_to_texto


class AutomatizacionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización de Documentos por Partidas")
        self.root.geometry("800x650")
        
        # Inicializar componentes
        self.excel_reader = ExcelReader()
        self.xml_processor = XMLProcessor()
        self.document_generator = DocumentGenerator()
        self.file_utils = FileUtils()
        
        # Crear la interfaz
        self.create_widgets()
    
    def create_widgets(self):
        # Configurar grid
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, minsize=120)
        
        # Selección de archivo Excel
        tk.Label(self.root, text="Archivo Excel de Partidas:", anchor='w').grid(row=0, column=0, padx=10, pady=5, sticky='ew')
        self.entry_excel_path = tk.Entry(self.root, width=50)
        self.entry_excel_path.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=self.select_excel_file).grid(row=0, column=2, padx=10, pady=5)
        
        # Separador
        ttk.Separator(self.root, orient='horizontal').grid(row=1, column=0, columnspan=3, sticky='ew', pady=10)
        
        # Sección de información común
        tk.Label(self.root, text="Información Común", font=('Helvetica', 10, 'bold')).grid(row=2, column=0, columnspan=3, sticky='w', padx=5)
        
        # Número de mensaje
        tk.Label(self.root, text="Número de mensaje de asignación:", anchor='w').grid(row=3, column=0, padx=10, pady=5, sticky='ew')
        self.entry_numero_mensaje = tk.Entry(self.root)
        self.entry_numero_mensaje.grid(row=3, column=1, padx=10, pady=5, sticky='ew')
        
        # Fecha del mensaje
        tk.Label(self.root, text="Fecha del mensaje de asignación:", anchor='w').grid(row=4, column=0, padx=10, pady=5, sticky='ew')
        self.entry_fecha_mensaje = tk.Entry(self.root)
        self.entry_fecha_mensaje.grid(row=4, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=lambda: self.select_date(self.entry_fecha_mensaje)).grid(row=4, column=2, padx=10, pady=5)
        
        # Fecha del documento
        tk.Label(self.root, text="Fecha de elaboración del documento:", anchor='w').grid(row=5, column=0, padx=10, pady=5, sticky='ew')
        self.entry_fecha_documento = tk.Entry(self.root)
        self.entry_fecha_documento.grid(row=5, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=lambda: self.select_date(self.entry_fecha_documento)).grid(row=5, column=2, padx=10, pady=5)
        
        # Número de oficio
        tk.Label(self.root, text="Número del Oficio de remisión:", anchor='w').grid(row=6, column=0, padx=10, pady=5, sticky='ew')
        self.entry_numero_oficio = tk.Entry(self.root)
        self.entry_numero_oficio.grid(row=6, column=1, padx=10, pady=5, sticky='ew')
        
        # Fecha de remisión
        tk.Label(self.root, text="Fecha del Oficio de Remisión:", anchor='w').grid(row=7, column=0, padx=10, pady=5, sticky='ew')
        self.entry_fecha_remision = tk.Entry(self.root)
        self.entry_fecha_remision.grid(row=7, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=lambda: self.select_date(self.entry_fecha_remision)).grid(row=7, column=2, padx=10, pady=5)
        
        # Mes asignado
        tk.Label(self.root, text="Mes asignado:", anchor='w').grid(row=8, column=0, padx=10, pady=5, sticky='ew')
        self.mes_asignado_var = tk.StringVar(self.root)
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        self.mes_asignado_var.set(meses[datetime.now().month - 1])  # Mes actual como predeterminado
        option_menu_meses = tk.OptionMenu(self.root, self.mes_asignado_var, *meses)
        option_menu_meses.grid(row=8, column=1, padx=10, pady=5, sticky='ew')
        
        # Separador
        ttk.Separator(self.root, orient='horizontal').grid(row=9, column=0, columnspan=3, sticky='ew', pady=10)
        
        # Barra de progreso
        tk.Label(self.root, text="Progreso:", anchor='w').grid(row=10, column=0, padx=10, pady=5, sticky='ew')
        self.progress_bar = ttk.Progressbar(self.root, length=400, mode='determinate')
        self.progress_bar.grid(row=10, column=1, columnspan=2, padx=10, pady=5, sticky='ew')
        
        # Botón de procesamiento
        tk.Button(self.root, text="Procesar", command=self.process_data, bg='#4CAF50', fg='white', height=2).grid(row=11, column=0, columnspan=3, pady=20, sticky='ew', padx=20)
        
        # Registro de actividad
        tk.Label(self.root, text="Registro de Actividad:", anchor='w').grid(row=12, column=0, columnspan=3, sticky='w', padx=10, pady=5)
        self.status_text = tk.Text(self.root, height=12, width=70)
        self.status_text.grid(row=13, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        
        # Agregar scrollbar
        scrollbar = tk.Scrollbar(self.root, command=self.status_text.yview)
        scrollbar.grid(row=13, column=3, sticky='ns')
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        # Configurar fila para expandir texto de estado
        self.root.rowconfigure(13, weight=1)
    
    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            self.entry_excel_path.delete(0, tk.END)
            self.entry_excel_path.insert(0, file_path)
    
    def select_date(self, entry_widget):
        DateSelector(self.root, entry_widget)
    
    def update_status(self, message):
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)  # Auto-scroll al final
        self.root.update_idletasks()
    
    def update_progress(self, value, maximum=100):
        self.progress_bar['value'] = (value / maximum) * 100
        self.root.update_idletasks()
    
    def process_data(self):
        # Obtener parámetros de la interfaz
        excel_path = self.entry_excel_path.get()
        numero_mensaje = self.entry_numero_mensaje.get()
        fecha_mensaje = self.entry_fecha_mensaje.get()
        fecha_documento = self.entry_fecha_documento.get()
        numero_oficio = self.entry_numero_oficio.get()
        fecha_remision = self.entry_fecha_remision.get()
        mes_asignado = self.mes_asignado_var.get()
        
        # Validar campos requeridos
        if not excel_path:
            messagebox.showerror("Error", "Por favor, seleccione un archivo Excel.")
            return
        
        if not numero_mensaje or not fecha_mensaje or not fecha_documento or not numero_oficio or not fecha_remision:
            messagebox.showerror("Error", "Todos los campos son obligatorios.")
            return
        
        # Convertir fechas a formato de texto
        try:
            fecha_mensaje_texto = convert_fecha_to_texto(fecha_mensaje)
            fecha_documento_texto = convert_fecha_to_texto(fecha_documento)
            fecha_remision_texto = convert_fecha_to_texto(fecha_remision)
        except ValueError:
            messagebox.showerror("Error", "El formato de fecha debe ser YYYY-MM-DD")
            return
        
        # Limpiar texto de estado
        self.status_text.delete(1.0, tk.END)
        
        # Procesar todo en el hilo principal
        self.process_on_main_thread(
            excel_path, numero_mensaje, fecha_mensaje, fecha_mensaje_texto,
            fecha_documento_texto, numero_oficio, fecha_remision_texto, mes_asignado
        )
    
    def process_on_main_thread(self, excel_path, numero_mensaje, fecha_mensaje_raw, fecha_mensaje,
                              fecha_documento, numero_oficio, fecha_remision, mes_asignado):
        try:
            self.update_status("Iniciando procesamiento...")
            
            # Leer archivo Excel
            self.update_status("Leyendo archivo Excel...")
            partidas = self.excel_reader.read_partidas(excel_path)
            self.update_status(f"Se encontraron {len(partidas)} partidas en el archivo.")
            
            # Directorio base (mismo que el archivo Excel)
            base_dir = os.path.dirname(excel_path)
            
            # Procesar cada partida
            total_partidas = len(partidas)
            for i, partida in enumerate(partidas):
                self.update_status(f"Procesando partida {partida['numero']} - {partida['descripcion']}...")
                
                # Actualizar progreso (nivel de partida)
                self.update_progress(i, total_partidas)
                
                # Actualizar la interfaz para que responda durante el procesamiento
                self.root.update()
                
                # Buscar archivos XML para esta partida
                partida_dir = os.path.join(base_dir, partida['numero'])
                if not os.path.exists(partida_dir):
                    self.update_status(f"  AVISO: Directorio para partida {partida['numero']} no encontrado.")
                    continue
                
                xml_files = self.file_utils.find_xml_files(partida_dir)
                self.update_status(f"  Se encontraron {len(xml_files)} archivos XML para procesar.")
                
                # Procesar cada archivo XML
                for j, xml_file in enumerate(xml_files):
                    self.update_status(f"  Procesando archivo: {os.path.basename(xml_file)}...")
                    
                    # Actualizar la interfaz para mantenerla responsiva
                    self.root.update()
                    
                    try:
                        # Formatear monto
                        monto_formateado = "$ {:,.2f}".format(partida['monto'])
                        
                        # Obtener directorio de salida
                        output_dir = os.path.dirname(xml_file)
                        
                        # Leer y procesar el XML
                        xml_data = self.xml_processor.read_xml(
                            xml_file,
                            numero_mensaje,
                            fecha_mensaje_raw,
                            mes_asignado,
                            monto_formateado,
                            fecha_documento,
                            partida['numero'],
                            numero_oficio,
                            fecha_remision
                        )
                        
                        # Generar documentos
                        self.document_generator.generate_all_documents(
                            xml_data,
                            output_dir,
                            partida
                        )
                        
                        self.update_status(f"  ✅ Documentos generados para {os.path.basename(xml_file)}.")
                    except Exception as e:
                        self.update_status(f"  ❌ ERROR al procesar {os.path.basename(xml_file)}: {str(e)}")
                        
                    # Actualizar la interfaz después de cada archivo
                    self.root.update()
            
            # Actualización final del progreso
            self.update_progress(total_partidas, total_partidas)
            self.update_status("¡Procesamiento completado con éxito!")
            messagebox.showinfo("Éxito", "Procesamiento completado con éxito")
        except Exception as e:
            self.update_status(f"ERROR: {str(e)}")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AutomatizacionApp(root)
    root.mainloop()
class AutomatizacionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización de Documentos por Partidas")
        self.root.geometry("800x650")
        
        # Inicializar componentes
        self.excel_reader = ExcelReader()
        self.xml_processor = XMLProcessor()
        self.document_generator = DocumentGenerator()
        self.file_utils = FileUtils()
        
        # Crear la interfaz
        self.create_widgets()
    
    def create_widgets(self):
        # Configurar grid
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, minsize=120)
        
        # Selección de archivo Excel
        tk.Label(self.root, text="Archivo Excel de Partidas:", anchor='w').grid(row=0, column=0, padx=10, pady=5, sticky='ew')
        self.entry_excel_path = tk.Entry(self.root, width=50)
        self.entry_excel_path.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=self.select_excel_file).grid(row=0, column=2, padx=10, pady=5)
        
        # Separador
        ttk.Separator(self.root, orient='horizontal').grid(row=1, column=0, columnspan=3, sticky='ew', pady=10)
        
        # Sección de información común
        tk.Label(self.root, text="Información Común", font=('Helvetica', 10, 'bold')).grid(row=2, column=0, columnspan=3, sticky='w', padx=5)
        
        # Número de mensaje
        tk.Label(self.root, text="Número de mensaje de asignación:", anchor='w').grid(row=3, column=0, padx=10, pady=5, sticky='ew')
        self.entry_numero_mensaje = tk.Entry(self.root)
        self.entry_numero_mensaje.grid(row=3, column=1, padx=10, pady=5, sticky='ew')
        
        # Fecha del mensaje
        tk.Label(self.root, text="Fecha del mensaje de asignación:", anchor='w').grid(row=4, column=0, padx=10, pady=5, sticky='ew')
        self.entry_fecha_mensaje = tk.Entry(self.root)
        self.entry_fecha_mensaje.grid(row=4, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=lambda: self.select_date(self.entry_fecha_mensaje)).grid(row=4, column=2, padx=10, pady=5)
        
        # Fecha del documento
        tk.Label(self.root, text="Fecha de elaboración del documento:", anchor='w').grid(row=5, column=0, padx=10, pady=5, sticky='ew')
        self.entry_fecha_documento = tk.Entry(self.root)
        self.entry_fecha_documento.grid(row=5, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=lambda: self.select_date(self.entry_fecha_documento)).grid(row=5, column=2, padx=10, pady=5)
        
        # Número de oficio
        tk.Label(self.root, text="Número del Oficio de remisión:", anchor='w').grid(row=6, column=0, padx=10, pady=5, sticky='ew')
        self.entry_numero_oficio = tk.Entry(self.root)
        self.entry_numero_oficio.grid(row=6, column=1, padx=10, pady=5, sticky='ew')
        
        # Fecha de remisión
        tk.Label(self.root, text="Fecha del Oficio de Remisión:", anchor='w').grid(row=7, column=0, padx=10, pady=5, sticky='ew')
        self.entry_fecha_remision = tk.Entry(self.root)
        self.entry_fecha_remision.grid(row=7, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=lambda: self.select_date(self.entry_fecha_remision)).grid(row=7, column=2, padx=10, pady=5)
        
        # Mes asignado
        tk.Label(self.root, text="Mes asignado:", anchor='w').grid(row=8, column=0, padx=10, pady=5, sticky='ew')
        self.mes_asignado_var = tk.StringVar(self.root)
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        self.mes_asignado_var.set(meses[datetime.now().month - 1])  # Mes actual como predeterminado
        option_menu_meses = tk.OptionMenu(self.root, self.mes_asignado_var, *meses)
        option_menu_meses.grid(row=8, column=1, padx=10, pady=5, sticky='ew')
        
        # Separador
        ttk.Separator(self.root, orient='horizontal').grid(row=9, column=0, columnspan=3, sticky='ew', pady=10)
        
        # Barra de progreso
        tk.Label(self.root, text="Progreso:", anchor='w').grid(row=10, column=0, padx=10, pady=5, sticky='ew')
        self.progress_bar = ttk.Progressbar(self.root, length=400, mode='determinate')
        self.progress_bar.grid(row=10, column=1, columnspan=2, padx=10, pady=5, sticky='ew')
        
        # Botón de procesamiento
        tk.Button(self.root, text="Procesar", command=self.process_data, bg='#4CAF50', fg='white', height=2).grid(row=11, column=0, columnspan=3, pady=20, sticky='ew', padx=20)
        
        # Registro de actividad
        tk.Label(self.root, text="Registro de Actividad:", anchor='w').grid(row=12, column=0, columnspan=3, sticky='w', padx=10, pady=5)
        self.status_text = tk.Text(self.root, height=12, width=70)
        self.status_text.grid(row=13, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        
        # Agregar scrollbar
        scrollbar = tk.Scrollbar(self.root, command=self.status_text.yview)
        scrollbar.grid(row=13, column=3, sticky='ns')
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        # Configurar fila para expandir texto de estado
        self.root.rowconfigure(13, weight=1)
    
    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            self.entry_excel_path.delete(0, tk.END)
            self.entry_excel_path.insert(0, file_path)
    
    def select_date(self, entry_widget):
        DateSelector(self.root, entry_widget)
    
    def update_status(self, message):
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)  # Auto-scroll al final
        self.root.update_idletasks()
    
    def update_progress(self, value, maximum=100):
        self.progress_bar['value'] = (value / maximum) * 100
        self.root.update_idletasks()
    
    def process_data(self):
        # Obtener parámetros de la interfaz
        excel_path = self.entry_excel_path.get()
        numero_mensaje = self.entry_numero_mensaje.get()
        fecha_mensaje = self.entry_fecha_mensaje.get()
        fecha_documento = self.entry_fecha_documento.get()
        numero_oficio = self.entry_numero_oficio.get()
        fecha_remision = self.entry_fecha_remision.get()
        mes_asignado = self.mes_asignado_var.get()
        
        # Validar campos requeridos
        if not excel_path:
            messagebox.showerror("Error", "Por favor, seleccione un archivo Excel.")
            return
        
        if not numero_mensaje or not fecha_mensaje or not fecha_documento or not numero_oficio or not fecha_remision:
            messagebox.showerror("Error", "Todos los campos son obligatorios.")
            return
        
        # Convertir fechas a formato de texto
        try:
            fecha_mensaje_texto = convert_fecha_to_texto(fecha_mensaje)
            fecha_documento_texto = convert_fecha_to_texto(fecha_documento)
            fecha_remision_texto = convert_fecha_to_texto(fecha_remision)
        except ValueError:
            messagebox.showerror("Error", "El formato de fecha debe ser YYYY-MM-DD")
            return
        
        # Limpiar texto de estado
        self.status_text.delete(1.0, tk.END)
        
        # Iniciar procesamiento en un hilo separado
        thread = threading.Thread(target=self.process_thread, args=(
            excel_path, numero_mensaje, fecha_mensaje, fecha_mensaje_texto,
            fecha_documento_texto, numero_oficio, fecha_remision_texto, mes_asignado
        ))
        thread.daemon = True
        thread.start()
    
    def process_thread(self, excel_path, numero_mensaje, fecha_mensaje_raw, fecha_mensaje,fecha_documento, numero_oficio, fecha_remision, mes_asignado):
        try:
            self.update_status("Iniciando procesamiento...")
            
            # Leer archivo Excel
            self.update_status("Leyendo archivo Excel...")
            partidas = self.excel_reader.read_partidas(excel_path)
            self.update_status(f"Se encontraron {len(partidas)} partidas en el archivo.")
            
            # Directorio base (mismo que el archivo Excel)
            base_dir = os.path.dirname(excel_path)
            
            # Procesar cada partida
            total_partidas = len(partidas)
            for i, partida in enumerate(partidas):
                self.update_status(f"Procesando partida {partida['numero']} - {partida['descripcion']}...")
                
                # Actualizar progreso (nivel de partida)
                self.update_progress(i, total_partidas)
                
                # Buscar archivos XML para esta partida
                partida_dir = os.path.join(base_dir, partida['numero'])
                if not os.path.exists(partida_dir):
                    self.update_status(f"  AVISO: Directorio para partida {partida['numero']} no encontrado.")
                    continue
                
                xml_files = self.file_utils.find_xml_files(partida_dir)
                self.update_status(f"  Se encontraron {len(xml_files)} archivos XML para procesar.")
                
                # Procesar cada archivo XML
                for j, xml_file in enumerate(xml_files):
                    self.update_status(f"  Procesando archivo: {os.path.basename(xml_file)}...")
                    
                    try:
                        # Formatear monto
                        monto_formateado = "$ {:,.2f}".format(partida['monto'])
                        
                        # Obtener directorio de salida
                        output_dir = os.path.dirname(xml_file)
                        
                        # Leer y procesar el XML
                        xml_data = self.xml_processor.read_xml(
                            xml_file,
                            numero_mensaje,
                            fecha_mensaje_raw,
                            mes_asignado,
                            monto_formateado,
                            fecha_documento,
                            partida['numero'],
                            numero_oficio,
                            fecha_remision
                        )
                        
                        # Generar documentos
                        self.document_generator.generate_all_documents(
                            xml_data,
                            output_dir,
                            partida
                        )
                        
                        self.update_status(f"  ✅ Documentos generados para {os.path.basename(xml_file)}.")
                    except Exception as e:
                        self.update_status(f"  ❌ ERROR al procesar {os.path.basename(xml_file)}: {str(e)}")
            
            # Actualización final del progreso
            self.update_progress(total_partidas, total_partidas)
            self.update_status("¡Procesamiento completado con éxito!")
            messagebox.showinfo("Éxito", "Procesamiento completado con éxito")
        except Exception as e:
            self.update_status(f"ERROR: {str(e)}")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AutomatizacionApp(root)
    root.mainloop()
