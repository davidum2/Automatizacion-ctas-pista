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


class AutomatizacionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización de Documentos por Partidas")
        self.root.geometry("800x700")  # Ajustamos altura ya que eliminaremos campos
        
        # Inicializar componentes
        self.excel_reader = ExcelReader()
        self.xml_processor = XMLProcessor()
        self.document_generator = DocumentGenerator()
        self.file_utils = FileUtils()
        
        # Listas de personal predefinidas
        self.lista_personal_recibe = [
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
        
        self.lista_personal_visto_bueno = [
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
        
        # Fecha del documento
        tk.Label(self.root, text="Fecha de elaboración del documento:", anchor='w').grid(row=3, column=0, padx=10, pady=5, sticky='ew')
        self.entry_fecha_documento = tk.Entry(self.root)
        self.entry_fecha_documento.grid(row=3, column=1, padx=10, pady=5, sticky='ew')
        tk.Button(self.root, text="Seleccionar", command=lambda: self.select_date(self.entry_fecha_documento)).grid(row=3, column=2, padx=10, pady=5)
        
        # Mes asignado
        tk.Label(self.root, text="Mes asignado:", anchor='w').grid(row=4, column=0, padx=10, pady=5, sticky='ew')
        self.mes_asignado_var = tk.StringVar(self.root)
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        self.mes_asignado_var.set(meses[datetime.now().month - 1])  # Mes actual como predeterminado
        option_menu_meses = tk.OptionMenu(self.root, self.mes_asignado_var, *meses)
        option_menu_meses.grid(row=4, column=1, padx=10, pady=5, sticky='ew')
        
        # Separador para sección de personal
        ttk.Separator(self.root, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky='ew', pady=10)
        
        # Sección de Personal
        tk.Label(self.root, text="Información de Personal", font=('Helvetica', 10, 'bold')).grid(row=6, column=0, columnspan=3, sticky='w', padx=5)
        
        # Persona que recibió la compra
        tk.Label(self.root, text="Persona que recibió la compra:", anchor='w').grid(row=7, column=0, padx=10, pady=5, sticky='ew')
        
        # Crear combobox para personal que recibió
        self.personal_recibio_var = tk.StringVar(self.root)
        
        # Preparar opciones para mostrar en el combobox
        opciones_recibio = []
        for persona in self.lista_personal_recibe:
            etiqueta = f"{persona['Grado_recibio_la_compra']} - {persona['Nombre_recibio_la_compra']} ({persona['Matricula_recibio_la_compra']})"
            opciones_recibio.append(etiqueta)
        
        # Combobox para persona que recibió
        self.combo_personal_recibio = ttk.Combobox(self.root, textvariable=self.personal_recibio_var, values=opciones_recibio, width=60)
        self.combo_personal_recibio.grid(row=7, column=1, padx=10, pady=5, sticky='ew')
        
        # Persona que dio el Vo. Bo.
        tk.Label(self.root, text="Persona que dio el Vo. Bo.:", anchor='w').grid(row=8, column=0, padx=10, pady=5, sticky='ew')
        
        # Crear combobox para personal de visto bueno
        self.personal_vobo_var = tk.StringVar(self.root)
        
        # Preparar opciones para mostrar en el combobox
        opciones_vobo = []
        for persona in self.lista_personal_visto_bueno:
            etiqueta = f"{persona['Grado_Vo_Bo']} - {persona['Nombre_Vo_Bo']} ({persona['Matricula_Vo_Bo']})"
            opciones_vobo.append(etiqueta)
        
        # Combobox para persona que dio visto bueno
        self.combo_personal_vobo = ttk.Combobox(self.root, textvariable=self.personal_vobo_var, values=opciones_vobo, width=60)
        self.combo_personal_vobo.grid(row=8, column=1, padx=10, pady=5, sticky='ew')
        
        # Separador antes de barra de progreso
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
    
    def obtener_datos_personal_recibio(self):
        """Obtiene los datos completos de la persona que recibió seleccionada"""
        if not self.personal_recibio_var.get():
            return None
            
        # Buscar la persona seleccionada en la lista
        etiqueta_seleccionada = self.personal_recibio_var.get()
        
        for persona in self.lista_personal_recibe:
            etiqueta = f"{persona['Grado_recibio_la_compra']} - {persona['Nombre_recibio_la_compra']} ({persona['Matricula_recibio_la_compra']})"
            if etiqueta == etiqueta_seleccionada:
                return persona
        
        return None
    
    def obtener_datos_personal_vobo(self):
        """Obtiene los datos completos de la persona que dio el visto bueno seleccionada"""
        if not self.personal_vobo_var.get():
            return None
            
        # Buscar la persona seleccionada en la lista
        etiqueta_seleccionada = self.personal_vobo_var.get()
        
        for persona in self.lista_personal_visto_bueno:
            etiqueta = f"{persona['Grado_Vo_Bo']} - {persona['Nombre_Vo_Bo']} ({persona['Matricula_Vo_Bo']})"
            if etiqueta == etiqueta_seleccionada:
                return persona
        
        return None
    
    def select_excel_file(self):
        """Abre el diálogo para seleccionar un archivo Excel"""
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if file_path:
            self.entry_excel_path.delete(0, tk.END)
            self.entry_excel_path.insert(0, file_path)
    
    def select_date(self, entry_widget):
        """Abre el selector de fecha para un widget de entrada"""
        DateSelector(self.root, entry_widget)
    
    def update_status(self, message):
        """Actualiza el registro de actividad con un nuevo mensaje"""
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)  # Auto-scroll al final
        self.root.update_idletasks()
    
    def update_progress(self, value, maximum=100):
        """Actualiza la barra de progreso"""
        self.progress_bar['value'] = (value / maximum) * 100
        self.root.update_idletasks()
    
    def process_data(self):
        # Obtener parámetros de la interfaz
        excel_path = self.entry_excel_path.get()
        fecha_documento = self.entry_fecha_documento.get()
        mes_asignado = self.mes_asignado_var.get()
        
        # Obtener información del personal seleccionado
        personal_recibio = self.obtener_datos_personal_recibio()
        personal_vobo = self.obtener_datos_personal_vobo()
        
        # Validar campos requeridos
        if not excel_path:
            messagebox.showerror("Error", "Por favor, seleccione un archivo Excel.")
            return
        
        if not fecha_documento:
            messagebox.showerror("Error", "La fecha de elaboración del documento es obligatoria.")
            return
        
        if not personal_recibio or not personal_vobo:
            messagebox.showerror("Error", "Por favor, seleccione el personal que recibió la compra y el que dio el visto bueno.")
            return
        
        # Convertir fecha a formato de texto
        try:
            fecha_documento_texto = convert_fecha_to_texto(fecha_documento)
        except ValueError:
            messagebox.showerror("Error", "El formato de fecha debe ser YYYY-MM-DD")
            return
        
        # Limpiar texto de estado
        self.status_text.delete(1.0, tk.END)
        
        # Procesar en hilo principal (para evitar problemas con la interfaz gráfica)
        self.process_on_main_thread(
            excel_path,
            fecha_documento_texto,
            mes_asignado,
            personal_recibio,
            personal_vobo
        )
    
    def process_on_main_thread(self, excel_path, fecha_documento, mes_asignado, 
                                personal_recibio, personal_vobo):
        try:
            self.update_status("Iniciando procesamiento...")
            
            # 1. Leer el archivo Excel con el nuevo formato
            self.update_status("Leyendo archivo Excel...")
            try:
                partidas = self.excel_reader.read_partidas(excel_path)
                self.update_status(f"Se encontraron {len(partidas)} partidas en el archivo.")
            except Exception as e:
                self.update_status(f"Error al leer el archivo Excel: {str(e)}")
                messagebox.showerror("Error", f"Error al leer el archivo Excel: {str(e)}")
                return
            
            # Directorio base (mismo que el archivo Excel)
            base_dir = os.path.dirname(excel_path)
            
            # Valores fijos para simplificar (estos ya no son usados desde la interfaz)
            numero_mensaje = "CC-001"  # Valor fijo para todos los documentos
            fecha_mensaje_raw = "2025-01-01"  # Fecha fija como ejemplo
            numero_oficio = "OF-001"  # Valor fijo para todos los documentos
            fecha_remision = "7 de marzo de 2025"  # Valor fijo formateado
            
            # Variables para seguimiento del progreso total
            total_facturas_procesadas = 0
            total_facturas = 0
            
            # Primero, contar todas las facturas para tener una idea del progreso total
            for partida in partidas:
                partida_dir = os.path.join(base_dir, partida['numero'])
                if os.path.exists(partida_dir):
                    # Contar subdirectorios (facturas potenciales)
                    subdirs = [d for d in os.listdir(partida_dir) 
                            if os.path.isdir(os.path.join(partida_dir, d))]
                    total_facturas += len(subdirs)
            
            self.update_status(f"Total de facturas a procesar: {total_facturas}")
            
            # 2. Procesar cada partida
            for i, partida in enumerate(partidas):
                self.update_status(f"\nProcesando partida {i+1}/{len(partidas)}: {partida['numero']} - {partida['descripcion']}...")
                
                # Actualizar progreso de partidas
                self.update_progress(i, len(partidas))
                
                # Buscar directorio de la partida
                partida_dir = os.path.join(base_dir, partida['numero'])
                if not os.path.exists(partida_dir):
                    self.update_status(f"  ⚠️ AVISO: Directorio para partida {partida['numero']} no encontrado.")
                    continue
                
                # 3. Escanear subdirectorios de facturas para esta partida
                subdirs = [d for d in os.listdir(partida_dir) 
                        if os.path.isdir(os.path.join(partida_dir, d))]
                
                self.update_status(f"  📂 Partida {partida['numero']}: {len(subdirs)} facturas encontradas.")
                
                # 4. Procesar cada factura en esta partida
                for j, subdir in enumerate(subdirs):
                    factura_dir = os.path.join(partida_dir, subdir)
                    self.update_status(f"  📄 Procesando factura {j+1}/{len(subdirs)} en {subdir}")
                    
                    # Actualizar la interfaz para mantenerla responsiva
                    self.root.update()
                    
                    # 5. Buscar archivos XML en este directorio de factura
                    xml_files = [f for f in os.listdir(factura_dir) 
                                if f.lower().endswith('.xml')]
                    
                    if not xml_files:
                        self.update_status(f"    ⚠️ No se encontró archivo XML en {subdir}")
                        continue
                    
                    # Tomar el primer XML encontrado
                    xml_file = os.path.join(factura_dir, xml_files[0])
                    
                    try:
                        # 6. Procesar la factura individual
                        self.update_status(f"    🔍 Analizando XML: {os.path.basename(xml_file)}...")
                        
                        # Formatear monto
                        monto_formateado = "$ {:,.2f}".format(partida['monto'])
                        
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
                        
                        # Agregar información del personal seleccionado
                        for key, value in personal_recibio.items():
                            xml_data[key] = value
                            
                        for key, value in personal_vobo.items():
                            xml_data[key] = value
                        
                        # Agregar el número adicional de la partida si existe
                        if 'numero_adicional' in partida and partida['numero_adicional']:
                            xml_data['numero_adicional'] = partida['numero_adicional']
                        
                        # 7. Generar documentos para esta factura
                        self.update_status(f"    📝 Generando documentos...")
                        self.document_generator.generate_all_documents(
                            xml_data,
                            factura_dir,
                            partida
                        )
                        
                        self.update_status(f"    ✅ Documentos generados para {os.path.basename(xml_file)}.")
                        total_facturas_procesadas += 1
                        
                        # Actualizar barra de progreso con el total de facturas
                        self.update_progress(total_facturas_procesadas, total_facturas)
                        
                    except Exception as e:
                        self.update_status(f"    ❌ ERROR al procesar {os.path.basename(xml_file)}: {str(e)}")
            
            # Actualización final del progreso
            self.update_progress(100, 100)
            
            # Resumen final
            self.update_status("\n===== RESUMEN DEL PROCESAMIENTO =====")
            self.update_status(f"Total de partidas encontradas: {len(partidas)}")
            self.update_status(f"Total de facturas procesadas: {total_facturas_procesadas}/{total_facturas}")
            self.update_status("¡Procesamiento completado con éxito!")
            
            messagebox.showinfo("Éxito", f"Procesamiento completado con éxito.\nSe procesaron {total_facturas_procesadas} facturas de {len(partidas)} partidas.")
            
        except Exception as e:
            self.update_status(f"ERROR GENERAL: {str(e)}")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
                               
            try:
                self.update_status("Iniciando procesamiento...")
                
                # 1. Leer todo el archivo Excel al principio
                self.update_status("Leyendo archivo Excel...")
                partidas = self.excel_reader.read_partidas(excel_path)
                self.update_status(f"Se encontraron {len(partidas)} partidas en el archivo.")
                
                # Directorio base (mismo que el archivo Excel)
                base_dir = os.path.dirname(excel_path)
                
                # Valores fijos para simplificar (estos ya no son usados desde la interfaz)
                numero_mensaje = "CC-001"  # Valor fijo para todos los documentos
                fecha_mensaje_raw = "2025-01-01"  # Fecha fija como ejemplo
                numero_oficio = "OF-001"  # Valor fijo para todos los documentos
                fecha_remision = "7 de marzo de 2025"  # Valor fijo formateado
                
                # Variables para seguimiento del progreso total
                total_facturas_procesadas = 0
                total_facturas = 0
                
                # Primero, contar todas las facturas para tener una idea del progreso total
                for partida in partidas:
                    partida_dir = os.path.join(base_dir, partida['numero'])
                    if os.path.exists(partida_dir):
                        # Contar subdirectorios (facturas potenciales)
                        subdirs = [d for d in os.listdir(partida_dir) 
                                if os.path.isdir(os.path.join(partida_dir, d))]
                        total_facturas += len(subdirs)
                
                self.update_status(f"Total de facturas a procesar: {total_facturas}")
                
                # 2. Procesar cada partida
                for i, partida in enumerate(partidas):
                    self.update_status(f"\nProcesando partida {i+1}/{len(partidas)}: {partida['numero']} - {partida['descripcion']}...")
                    
                    # Actualizar progreso de partidas
                    self.update_progress(i, len(partidas))
                    
                    # Buscar directorio de la partida
                    partida_dir = os.path.join(base_dir, partida['numero'])
                    if not os.path.exists(partida_dir):
                        self.update_status(f"  ⚠️ AVISO: Directorio para partida {partida['numero']} no encontrado.")
                        continue
                    
                    # 3. Escanear subdirectorios de facturas para esta partida
                    subdirs = [d for d in os.listdir(partida_dir) 
                            if os.path.isdir(os.path.join(partida_dir, d))]
                    
                    self.update_status(f"  📂 Partida {partida['numero']}: {len(subdirs)} facturas encontradas.")
                    
                    # 4. Procesar cada factura en esta partida
                    for j, subdir in enumerate(subdirs):
                        factura_dir = os.path.join(partida_dir, subdir)
                        self.update_status(f"  📄 Procesando factura {j+1}/{len(subdirs)} en {subdir}")
                        
                        # Actualizar la interfaz para mantenerla responsiva
                        self.root.update()
                        
                        # 5. Buscar archivos XML en este directorio de factura
                        xml_files = [f for f in os.listdir(factura_dir) 
                                    if f.lower().endswith('.xml')]
                        
                        if not xml_files:
                            self.update_status(f"    ⚠️ No se encontró archivo XML en {subdir}")
                            continue
                        
                        # Tomar el primer XML encontrado
                        xml_file = os.path.join(factura_dir, xml_files[0])
                        
                        try:
                            # 6. Procesar la factura individual
                            self.update_status(f"    🔍 Analizando XML: {os.path.basename(xml_file)}...")
                            
                            # Formatear monto
                            monto_formateado = "$ {:,.2f}".format(partida['monto'])
                            
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
                            
                            # Agregar información del personal seleccionado
                            for key, value in personal_recibio.items():
                                xml_data[key] = value
                                
                            for key, value in personal_vobo.items():
                                xml_data[key] = value
                            
                            # 7. Generar documentos para esta factura
                            self.update_status(f"    📝 Generando documentos...")
                            self.document_generator.generate_all_documents(
                                xml_data,
                                factura_dir,
                                partida
                            )
                            
                            self.update_status(f"    ✅ Documentos generados para {os.path.basename(xml_file)}.")
                            total_facturas_procesadas += 1
                            
                            # Actualizar barra de progreso con el total de facturas
                            self.update_progress(total_facturas_procesadas, total_facturas)
                            
                        except Exception as e:
                            self.update_status(f"    ❌ ERROR al procesar {os.path.basename(xml_file)}: {str(e)}")
                
                # Actualización final del progreso
                self.update_progress(100, 100)
                
                # Resumen final
                self.update_status("\n===== RESUMEN DEL PROCESAMIENTO =====")
                self.update_status(f"Total de partidas encontradas: {len(partidas)}")
                self.update_status(f"Total de facturas procesadas: {total_facturas_procesadas}/{total_facturas}")
                self.update_status("¡Procesamiento completado con éxito!")
                
                messagebox.showinfo("Éxito", f"Procesamiento completado con éxito.\nSe procesaron {total_facturas_procesadas} facturas de {len(partidas)} partidas.")
                
            except Exception as e:
                self.update_status(f"ERROR GENERAL: {str(e)}")
                messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AutomatizacionApp(root)
    root.mainloop()