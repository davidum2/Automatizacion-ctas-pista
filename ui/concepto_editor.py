import tkinter as tk
from tkinter import simpledialog, messagebox
import re



def formatear_conceptos_automatico(conceptos_originales):
    """
    Formatea todos los conceptos en un formato unificado.

    Args:
        conceptos_originales (dict): Diccionario con los conceptos originales {descripcion: cantidad}

    Returns:
        str: Texto formateado de conceptos
    """
    # Crear una lista vacía para almacenar los conceptos formateados
    conceptos_texto = []
    
    # Ordenar los conceptos por cantidad (de mayor a menor)
    sorted_items = sorted(conceptos_originales.items(), key=lambda x: x[1], reverse=True)
    
    # Procesar cada concepto
    for descripcion, cantidad in sorted_items:
        # Limpiar descripción (eliminar numeración al inicio si existe)
        clean_desc = re.sub(r'^\d+\s*\.\s*', '', descripcion).strip()
        
        # Formatear cantidad con 3 decimales y añadir a la lista
        conceptos_texto.append(f"{cantidad:.3f} {clean_desc}")
    
    # Unir todos los conceptos con comas
    return ", ".join(conceptos_texto)


# aqui termina

class SimpleConceptoDialog(simpledialog.Dialog):
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
            messagebox.showwarning("Advertencia", "El texto no puede estar vacío")
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
