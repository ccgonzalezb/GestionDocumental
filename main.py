from model import PDFProcessor
from view import PDFExcelView
from controller import PDFController

if __name__ == "__main__":
    ruta_excel = "informacion_documento.xlsx"
    tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
    
    model = PDFProcessor(tesseract_cmd=tesseract_cmd, excel_path=ruta_excel)
    view = PDFExcelView(controller=None)
    controller = PDFController(model=model, view=view)
    view.controller = controller

    view.mainloop()
