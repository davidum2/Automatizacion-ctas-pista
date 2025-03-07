import tkinter as tk
from tkinter import Toplevel
from tkcalendar import Calendar
from datetime import datetime

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
        self.top = Toplevel(parent)
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
