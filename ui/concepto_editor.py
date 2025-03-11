import tkinter as tk
from tkinter import ttk, messagebox
import re

class ConceptoEditorWindow:
    """
    Ventana para editar y reorganizar los conceptos de una factura.
    Permite al usuario ver los conceptos extraídos del XML y reorganizarlos
    según los conceptos que maneja la propia partida.
    """

    def __init__(self, parent, conceptos_originales, partida_descripcion, callback):
        """
        Inicializa la ventana de edición de conceptos.

        Args:
            parent: Ventana padre
            conceptos_originales (dict): Diccionario con los conceptos originales {descripcion: cantidad}
            partida_descripcion (str): Descripción de la partida actual
            callback (function): Función a llamar cuando se confirman los cambios,
                                recibe el texto de conceptos editado como argumento
        """
        self.parent = parent
        self.conceptos_originales = conceptos_originales
        self.partida_descripcion = partida_descripcion
        self.callback = callback
        self.concepto_final = None

        # Crear ventana flotante
        self.top = tk.Toplevel(parent)
        self.top.title("Edición de Conceptos")
        self.top.geometry("800x600")
        self.top.transient(parent)  # Hacer la ventana dependiente de la ventana principal
        self.top.grab_set()  # Bloquear interacción con la ventana principal

        # Centrar la ventana
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (800 // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (600 // 2)
        self.top.geometry(f"800x600+{x}+{y}")

        # Crear la interfaz
        self.create_widgets()

        # Poblar los datos
        self.populate_data()

    def create_widgets(self):
        """Crea los widgets de la interfaz."""
        # Sección superior: información de la partida
        frame_info = tk.Frame(self.top, padx=10, pady=10)
        frame_info.pack(fill='x')

        tk.Label(frame_info, text="Partida:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w')
        tk.Label(frame_info, text=self.partida_descripcion, wraplength=700).grid(row=0, column=1, sticky='w')

        # Separador
        ttk.Separator(self.top, orient='horizontal').pack(fill='x', padx=10, pady=5)

        # Sección media: conceptos originales (tabla)
        frame_conceptos = tk.Frame(self.top, padx=10, pady=10)
        frame_conceptos.pack(fill='both', expand=True)

        tk.Label(frame_conceptos, text="Conceptos Originales:", font=('Arial', 10, 'bold')).pack(anchor='w')

        # Crear tabla con scrollbar
        frame_tabla = tk.Frame(frame_conceptos)
        frame_tabla.pack(fill='both', expand=True, pady=5)

        # Scrollbar vertical
        scrollbar_y = tk.Scrollbar(frame_tabla)
        scrollbar_y.pack(side='right', fill='y')

        # Scrollbar horizontal
        scrollbar_x = tk.Scrollbar(frame_tabla, orient='horizontal')
        scrollbar_x.pack(side='bottom', fill='x')

        # Tabla (Treeview)
        self.tabla = ttk.Treeview(
            frame_tabla,
            columns=("cantidad", "descripcion"),
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )

        # Configurar columnas
        self.tabla.column("#0", width=50, stretch=False)  # Columna de índice
        self.tabla.column("cantidad", width=100, stretch=False)
        self.tabla.column("descripcion", width=600)

        # Encabezados
        self.tabla.heading("#0", text="No.")
        self.tabla.heading("cantidad", text="Cantidad")
        self.tabla.heading("descripcion", text="Descripción")

        self.tabla.pack(fill='both', expand=True)

        # Configurar scrollbars
        scrollbar_y.config(command=self.tabla.yview)
        scrollbar_x.config(command=self.tabla.xview)

        # Separador
        ttk.Separator(self.top, orient='horizontal').pack(fill='x', padx=10, pady=5)

        # Sección inferior: conceptos editados
        frame_edicion = tk.Frame(self.top, padx=10, pady=10)
        frame_edicion.pack(fill='x')

        tk.Label(frame_edicion, text="Concepto Editado:", font=('Arial', 10, 'bold')).pack(anchor='w')

        # Campo de sugerencia (generado automáticamente)
        tk.Label(frame_edicion, text="Sugerencia:").pack(anchor='w', pady=(5, 0))
        self.sugerencia_text = tk.Text(frame_edicion, height=3, width=80, wrap='word')
        self.sugerencia_text.pack(fill='x', pady=5)
        self.sugerencia_text.config(state='disabled')  # Solo lectura

        # Botón para copiar sugerencia al campo de edición
        tk.Button(frame_edicion, text="Usar Sugerencia", command=self.usar_sugerencia).pack(anchor='e', pady=5)

        # Campo de edición
        tk.Label(frame_edicion, text="Editar Concepto Final:").pack(anchor='w', pady=(5, 0))
        self.concepto_text = tk.Text(frame_edicion, height=5, width=80, wrap='word')
        self.concepto_text.pack(fill='x', pady=5)

        # Botones de acción
        frame_botones = tk.Frame(self.top)
        frame_botones.pack(fill='x', padx=10, pady=10)

        tk.Button(frame_botones, text="Cancelar", command=self.top.destroy).pack(side='right', padx=5)
        tk.Button(frame_botones, text="Confirmar", command=self.confirmar, bg='#4CAF50', fg='white').pack(side='right', padx=5)

    def populate_data(self):
        """Puebla la tabla con los datos de conceptos."""
        # Limpiar tabla si ya tenía datos
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        # Llenar la tabla con los conceptos
        for i, (descripcion, cantidad) in enumerate(self.conceptos_originales.items(), 1):
            self.tabla.insert("", "end", text=str(i), values=(f"{cantidad:.3f}", descripcion))

        # Generar sugerencia automática
        self.generar_sugerencia()

    def generar_sugerencia(self):
        """Genera una sugerencia automática basada en los conceptos."""
        total_items = sum(self.conceptos_originales.values())

        # Decidir formato según los datos
        if len(self.conceptos_originales) == 1:
            # Si solo hay un concepto, usar ese directamente
            descripcion = list(self.conceptos_originales.keys())[0]
            cantidad = list(self.conceptos_originales.values())[0]
            sugerencia = f"{cantidad:.3f} {descripcion}"
        else:
            # Si hay múltiples conceptos, crear una lista compacta
            conceptos_texto = []
            for descripcion, cantidad in self.conceptos_originales.items():
                # Limpiar descripción (quitar números de inicio si existen)
                clean_desc = re.sub(r'^\d+\s*\.\s*', '', descripcion).strip()
                conceptos_texto.append(f"{cantidad:.3f} {clean_desc}")

            # Unir con comas
            sugerencia = ", ".join(conceptos_texto)

            # Si es demasiado largo, simplificar
            if len(sugerencia) > 200:
                sugerencia = f"{total_items:.3f} unidades de varios artículos"

        # Actualizar el widget de sugerencia
        self.sugerencia_text.config(state='normal')
        self.sugerencia_text.delete(1.0, tk.END)
        self.sugerencia_text.insert(tk.END, sugerencia)
        self.sugerencia_text.config(state='disabled')

    def usar_sugerencia(self):
        """Copia la sugerencia al campo de edición."""
        self.concepto_text.delete(1.0, tk.END)
        self.concepto_text.insert(tk.END, self.sugerencia_text.get(1.0, tk.END).strip())

    def confirmar(self):
        """Confirma los cambios y cierra la ventana."""
        concepto = self.concepto_text.get(1.0, tk.END).strip()

        if not concepto:
            messagebox.showerror("Error", "El concepto no puede estar vacío.")
            return

        # Llamar al callback con el texto editado
        self.callback(concepto)

        # Cerrar la ventana
        self.top.destroy()


# Ejemplo de uso:
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Prueba Editor de Conceptos")
    root.geometry("500x300")

    # Datos de ejemplo
    conceptos_ejemplo = {
        "Bolígrafo Punto Fino Negro": 12.0,
        "Lápiz HB con Goma": 24.0,
        "Cuaderno Profesional Raya 100 hojas": 10.0,
        "Folder Tamaño Carta Color Azul": 50.0
    }

    def on_edit_conceptos():
        def callback(concepto_editado):
            print(f"Concepto editado: {concepto_editado}")
            messagebox.showinfo("Concepto Editado", f"Concepto final:\n{concepto_editado}")

        ConceptoEditorWindow(root, conceptos_ejemplo, "24101 - MATERIALES Y ÚTILES DE OFICINA", callback)

    tk.Button(root, text="Editar Conceptos", command=on_edit_conceptos).pack(padx=20, pady=20)

    root.mainloop()
