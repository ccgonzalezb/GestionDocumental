import threading

class PDFController:
    def __init__(self, model, view):
        self.model = model
        self.view = view

    def procesar_pdf(self, ruta_pdf):
        self.model.procesar_pdf(ruta_pdf)

    def cargar_archivos(self):
        """Procesa los archivos seleccionados y actualiza la barra de progreso."""
        if not self.view.pdf_files:
            self.view.show_warning_message("Advertencia", "No se han seleccionado archivos para cargar.")
            return

        def procesar():
            total_files = len(self.view.pdf_files)
            for index, file_path in enumerate(self.view.pdf_files):
                # Calcula el número del documento actual y el total
                documento_actual = index + 1
                self.view.update_progress(documento_actual, total_files)  # Actualiza la barra de progreso.
                self.procesar_pdf(file_path)

            self.view.update_progress(total_files, total_files)  # Asegura que la barra llegue al 100%.
            self.view.show_info_message("Éxito", "Los archivos PDF han sido procesados y los datos se han guardado en el archivo Excel.")
            self.view.clear_table()  # Limpiar la tabla visual después de procesar.
            self.view.pdf_files.clear()  # Limpia la lista interna de archivos.

        # Ejecuta el procesamiento en un hilo separado para no bloquear la interfaz gráfica.
        thread = threading.Thread(target=procesar)
        thread.start()
