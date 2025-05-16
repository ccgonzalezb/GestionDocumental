# Procesador de Documentos PDF con OCR y Exportación a Excel

Este proyecto permite procesar archivos PDF para extraer información clave (como número de radicado, fechas y otros metadatos), incluso si los documentos contienen texto escaneado (imágenes). El programa utiliza OCR (Reconocimiento Óptico de Caracteres) mediante Tesseract y exporta los resultados a un archivo Excel.

## Funcionalidades

- Interfaz gráfica simple para cargar documentos PDF.
- Extracción automática de:
  - Número de radicado
  - Fechas relevantes
  - Origen, destino, tipo de documento, asunto y otros campos administrativos
- Aplicación de OCR en páginas sin texto usando Tesseract.
- Generación automática de archivo Excel con la información procesada.
- Barra de progreso visual.
- Tabla interactiva para visualizar los archivos seleccionados.

## Tecnologías y bibliotecas utilizadas

- Python 3.x
- [Tkinter](https://docs.python.org/3/library/tkinter.html) (Interfaz gráfica)
- [pdfplumber](https://github.com/jsvine/pdfplumber)
- [pytesseract](https://github.com/madmaze/pytesseract)
- [pdf2image](https://github.com/Belval/pdf2image)
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

## Requisitos

- Python 3.8 o superior
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) instalado y accesible desde el sistema.

### Instalación de dependencias (Windows)

1. Clona el repositorio:

```bash
git clone https://github.com/tu-usuario/tu-repo.git
cd tu-repo
```

2. Crea un entorno virtual (opcional pero recomendado):

```bash
python -m venv venv
venv\Scripts\activate
```

3. Instala los paquetes necesarios:

```bash
pip install -r requirements.txt
```

4. Asegúrate de tener Tesseract instalado. Puedes descargarlo desde:

[https://github.com/tesseract-ocr/tesseract/wiki](https://github.com/tesseract-ocr/tesseract/wiki)

Configura su ruta en el archivo `main.py`:

```python
tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
```

## Uso

1. Ejecuta la aplicación:

```bash
python main.py
```

2. En la interfaz:
   - Haz clic en "Nuevo Archivo" para seleccionar uno o varios archivos PDF.
   - Presiona "Carga" para iniciar el procesamiento.
   - Usa "Examinar" para abrir el archivo Excel generado.

3. El archivo Excel con la información extraída se guardará como:  
   `informacion_documento.xlsx`

## Estructura del Proyecto

```
├── controller.py
├── main.py
├── model.py
├── view.py
├── informacion_documento_2.xlsx  # Se genera automáticamente
└── README.md
```

## Autor 
Cristian Camilo González Blanco