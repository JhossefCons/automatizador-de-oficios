# Automatizador de Oficios

Este proyecto automatiza el procesamiento de documentos oficiales ("oficios") en formato PDF, extrayendo información clave como el número de serie y el asunto, y organizándola en hojas de cálculo Excel para facilitar su gestión y análisis.

## Funcionalidades

- **Extracción de texto de PDFs**: Soporta tanto PDFs con texto nativo como PDFs escaneados mediante OCR (Reconocimiento Óptico de Caracteres).
- **Interfaz gráfica intuitiva**: Basada en Tkinter para una experiencia de usuario sencilla.
- **Procesamiento por lotes**: Permite seleccionar y procesar múltiples archivos PDF simultáneamente.
- **Extracción automática de datos**: Identifica automáticamente el número de serie (formato 5-XXX/XXX) y el asunto del documento.
- **Generación de reportes**: Crea archivos Excel organizados por fecha con toda la información extraída.
- **Registro de actividad**: Muestra un log detallado del proceso de extracción para facilitar la depuración.

## Versiones del Programa

### 1. Procesador de Oficios (GUI con OCR)
- **Archivo**: `procesador_oficios.py`
- **Descripción**: Versión completa con interfaz gráfica que utiliza OCR para procesar PDFs escaneados.
- **Características**: Preprocesamiento de imágenes para mejorar la precisión del OCR, barra de progreso, apertura automática del Excel generado.

### 2. Extractor de Texto (CLI)
- **Archivo**: `extractor-texto.py`
- **Descripción**: Versión simplificada de línea de comandos para PDFs con texto nativo.
- **Características**: Procesamiento rápido sin OCR, ideal para documentos digitales.

## Requisitos del Sistema

- Python 3.7+
- Tesseract OCR instalado y configurado
- Poppler (para conversión de PDF a imágenes)

## Dependencias de Python

```
pandas
tkinter (incluido en Python estándar)
pytesseract
pdf2image
Pillow
PyPDF2
numpy
```

## Instalación

1. **Instalar Python**: Asegúrate de tener Python 3.7 o superior instalado.

2. **Instalar Tesseract OCR**:
   - Descarga e instala Tesseract desde: https://github.com/UB-Mannheim/tesseract/wiki
   - Asegúrate de que esté en el PATH del sistema o actualiza la ruta en el código.

3. **Instalar dependencias**:
   ```bash
   pip install pandas pytesseract pdf2image Pillow PyPDF2 numpy
   ```

4. **Instalar Poppler** (requerido por pdf2image):
   - En Windows: Descarga desde https://blog.alivate.com.au/poppler-windows/
   - En Linux/Mac: `sudo apt-get install poppler-utils` o `brew install poppler`

## Uso

### Versión GUI (Recomendada)
1. Ejecuta `procesador_oficios.py`:
   ```bash
   python procesador_oficios.py
   ```

2. Haz clic en "Cargar archivos PDF" para seleccionar los documentos.

3. Presiona "Procesar Archivos" para iniciar la extracción.

4. Una vez completado, usa "Abrir Excel" para ver los resultados.

### Versión CLI
1. Ejecuta `extractor-texto.py`:
   ```bash
   python extractor-texto.py
   ```

2. Selecciona los archivos PDF en el diálogo.

3. Los resultados se guardarán automáticamente en un archivo Excel.

## Estructura del Proyecto

```
automatizador-de-oficios/
├── extractor-texto.py          # Versión CLI simple
├── procesador_oficios.py       # Versión GUI con OCR
├── ProcesadorOficios.spec      # Especificación para PyInstaller
├── README.md                   # Este archivo
├── .gitignore                  # Archivos ignorados por Git
├── build/                      # Archivos de construcción (PyInstaller)
│   └── ProcesadorOficios/      # Ejecutable generado
└── Resultados/                 # Carpeta de salida (ignorada por Git)
    ├── debug_ocr_text.txt      # Texto extraído para depuración
    └── oficios_YYYY-MM-DD.xlsx # Resultados organizados por fecha
```

## Ejecutable Precompilado

En la carpeta `build/ProcesadorOficios/` se encuentra un ejecutable independiente generado con PyInstaller. Este archivo permite ejecutar el programa sin necesidad de instalar Python o las dependencias en el sistema destino.

Para regenerar el ejecutable:
```bash
pip install pyinstaller
pyinstaller ProcesadorOficios.spec
```

## Notas Técnicas

- **OCR**: Utiliza Tesseract con idioma español para mejorar la precisión en documentos en español.
- **Preprocesamiento**: Las imágenes se convierten a blanco y negro y se aumenta el contraste antes del OCR.
- **Formato de salida**: Los archivos Excel se nombran con el formato `oficios_YYYY-MM-DD.xlsx`.
- **Manejo de errores**: El programa registra errores y continúa procesando otros archivos si uno falla.

## Contribución

Si deseas contribuir al proyecto:
1. Haz un fork del repositorio.
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`).
3. Commit tus cambios (`git commit -am 'Agrega nueva funcionalidad'`).
4. Push a la rama (`git push origin feature/nueva-funcionalidad`).
5. Abre un Pull Request.

## Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo LICENSE para más detalles.

## Soporte

Para reportar bugs o solicitar nuevas funcionalidades, por favor abre un issue en el repositorio de GitHub.
