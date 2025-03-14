"""
Diálogos personalizados para la interfaz de usuario
"""
import tkinter as tk
from tkinter import simpledialog, messagebox
from tkcalendar import Calendar
from datetime import datetime
import re

class DateSelector:
    """
    Ventana emergente para seleccionar una fecha con un calendario.
    """
    
    def __init__(self, parent, entry_widget):
        """
        Inicializa el selector de fechas.
        
        Args:
            parent: Ventana padre
            entry_widget: Widget de entrada donde se colocará la fecha seleccionada
        """
        self.parent = parent
        self.entry_widget = entry_widget
        
        # Crear ventana flotante
        self.top = tk.Toplevel(parent)
        self.top.title("Seleccionar Fecha")
        self.top.transient(parent)  # Hacer la ventana dependiente de la ventana principal
        self.top.grab_set()  # Bloquear interacción con la ventana principal
        
        # Centrar la ventana
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (300 // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (300 // 2)
        self.top.geometry(f"300x300+{x}+{y}")
        
        # Crear calendario
        self.cal = Calendar(self.top, selectmode='day', date_pattern='yyyy-mm-dd')
        self.cal.pack(pady=20, expand=True, fill='both')
        
        # Obtener fecha actual del campo si existe
        current_value = entry_widget.get()
        if current_value:
            try:
                current_date = datetime.strptime(current_value, '%Y-%m-%d')
                self.cal.selection_set(current_date)
            except ValueError:
                pass
        
        # Botones
        btn_frame = tk.Frame(self.top)
        btn_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Button(btn_frame, text="Seleccionar", command=self.select_date, 
                 bg='#4CAF50', fg='white').pack(side='left', padx=10, expand=True)
        tk.Button(btn_frame, text="Cancelar", command=self.top.destroy).pack(side='right', padx=10, expand=True)
    
    def select_date(self):
        """
        Selecciona la fecha y la coloca en el widget de entrada.
        """
        selected_date = self.cal.selection_get()
        self.entry_widget.delete(0, tk.END)
        self.entry_widget.insert(0, selected_date.strftime('%Y-%m-%d'))
        self.top.destroy()


class ConceptoEditor(simpledialog.Dialog):
    """Diálogo para editar conceptos"""

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
            messagebox.showwarning("Advertencia", "El texto no puede estar vacío")
            return False
        return True

    def apply(self):
        """El resultado ya se guardó en validate()"""
        pass


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

def editar_conceptos(parent, conceptos_originales, partida_descripcion):
    """
    Muestra un diálogo para editar conceptos y devuelve el texto editado.
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

        # Mostrar el diálogo
        dialog = ConceptoEditor(parent, conceptos_originales, partida_descripcion)

        # Si se canceló o dio error, usar la sugerencia
        if not hasattr(dialog, 'result') or not dialog.result:
            return sugerencia

        return dialog.result

    except Exception as e:
        # Si hay cualquier error, devolver la sugerencia automática
        print(f"Error en editor de conceptos: {e}")
        return sugerencia