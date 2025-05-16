import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os

class PDFExcelView(tk.Tk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("Carga Documentos")
        self.geometry("600x400")

        self.pdf_files = []

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding=(10, 10, 10, 10))
        main_frame.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=10)

        ttk.Label(header_frame, text="Seleccione los archivos a cargar").pack(side=tk.LEFT)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        self.new_file_button = ttk.Button(button_frame, text="Nuevo Archivo", command=self.select_pdf)
        self.new_file_button.pack(side=tk.LEFT, padx=5)

        self.load_button = ttk.Button(button_frame, text="Carga", command=self.cargar_archivos)
        self.load_button.pack(side=tk.LEFT, padx=5)

        self.section_label = ttk.Label(button_frame, text="Consulta de Excel generado")
        self.section_label.pack(side=tk.LEFT, padx=5)

        self.browse_button = ttk.Button(button_frame, text="Examinar", command=self.abrir_excel)
        self.browse_button.pack(side=tk.LEFT, padx=5)

        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.pack(pady=10)

        self.progress_label = ttk.Label(main_frame, text="")
        self.progress_label.pack(pady=5)

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        columns = ("#", "Ruta", "Eliminar")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("#", text="N°")
        self.tree.heading("Ruta", text="Ruta")
        self.tree.heading("Eliminar", text="Eliminar")
        self.tree.column("#", width=30, anchor=tk.CENTER)
        self.tree.column("Ruta", width=450, anchor=tk.W)
        self.tree.column("Eliminar", width=80, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind('<Double-1>', self.on_double_click)

    def add_file_to_table(self, file_path):
        next_id = len(self.tree.get_children()) + 1
        self.tree.insert("", "end", values=(next_id, file_path, "Eliminar"))

    def remove_file_from_table(self, item):
        self.tree.delete(item)

    def clear_table(self):
        """Limpia todos los elementos de la tabla."""
        for item in self.tree.get_children():
            self.tree.delete(item)    

    def on_double_click(self, event):
        item = self.tree.selection()[0]
        self.remove_file_from_table(item)

    def show_info_message(self, title, message):
        messagebox.showinfo(title, message)

    def show_error_message(self, title, message):
        messagebox.showerror(title, message)

    def show_warning_message(self, title, message):
        messagebox.showwarning(title, message)

    def select_pdf(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if file_paths:
            for file_path in file_paths:
                if not self.is_file_in_list(file_path):
                    self.pdf_files.append(file_path)
                    self.add_file_to_table(file_path)
                else:
                    self.show_warning_message("Advertencia", f"El archivo {file_path} ya ha sido agregado.")

    def is_file_in_list(self, file_path):
        return file_path in self.pdf_files

    def cargar_archivos(self):
        if not self.pdf_files:
            self.show_warning_message("Advertencia", "No se han seleccionado archivos para cargar.")
            return
        self.controller.cargar_archivos()

    def abrir_excel(self):
        ruta_excel = self.controller.model.excel_path
        if os.path.exists(ruta_excel):
            os.startfile(ruta_excel)
        else:
            self.show_warning_message("Advertencia", f"No se encontró el archivo {ruta_excel}.")

    def update_progress(self, current, total):
        """Actualiza la barra de progreso y la etiqueta con el número de documento actual y el total."""
        progress_value = (current / total) * 100
        self.progress['value'] = progress_value
        self.progress_label.config(text=f"Procesando documento {current} de {total}")
        self.update_idletasks()  # Refresca la interfaz gráfica
