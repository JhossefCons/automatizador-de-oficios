import os
import re
import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from tkinter.constants import END
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import tempfile
import numpy as np
import subprocess
import platform

# Configuración para Tesseract OCR (ajustar según tu instalación)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Oficios PDF (OCR)")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_files = []
        self.results_folder = os.path.join(os.getcwd(), "Resultados")
        self.current_excel_file = ""
        
        # Crear carpeta de resultados si no existe
        if not os.path.exists(self.results_folder):
            os.makedirs(self.results_folder)
        
        # Crear interfaz
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = tk.Label(
            main_frame, 
            text="Procesador de Oficios PDF (OCR)", 
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Nota sobre OCR
        ocr_note = tk.Label(
            main_frame,
            text="Este programa utiliza OCR para procesar PDFs escaneados",
            font=("Arial", 10, "italic"),
            fg="#555555"
        )
        ocr_note.pack(pady=(0, 10))
        
        # Frame para botones superiores
        top_buttons_frame = tk.Frame(main_frame)
        top_buttons_frame.pack(fill=tk.X, pady=10)
        
        # Botón para cargar archivos
        load_button = tk.Button(
            top_buttons_frame,
            text="Cargar archivos PDF",
            command=self.load_files,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5
        )
        load_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón para abrir Excel
        self.open_excel_button = tk.Button(
            top_buttons_frame,
            text="Abrir Excel",
            command=self.open_excel_file,
            bg="#FF9800",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5,
            state=tk.DISABLED
        )
        self.open_excel_button.pack(side=tk.LEFT)
        
        # Etiqueta de archivos seleccionados
        files_label = tk.Label(
            main_frame,
            text="Archivos seleccionados:",
            font=("Arial", 10)
        )
        files_label.pack(anchor="w", pady=(20, 5))
        
        # Lista de archivos
        self.files_listbox = tk.Listbox(
            main_frame,
            height=6,
            selectmode=tk.SINGLE,
            font=("Arial", 9)
        )
        self.files_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Botón para eliminar selección
        remove_button = tk.Button(
            main_frame,
            text="Eliminar seleccionado",
            command=self.remove_selected,
            bg="#f44336",
            fg="white",
            font=("Arial", 9),
            padx=5,
            pady=2
        )
        remove_button.pack(anchor="e", pady=(0, 20))
        
        # Botón de procesamiento
        process_button = tk.Button(
            main_frame,
            text="Procesar Archivos",
            command=self.process_files,
            bg="#2196F3",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=15,
            pady=8
        )
        process_button.pack(pady=20)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(
            main_frame,
            orient="horizontal",
            length=400,
            mode="determinate"
        )
        self.progress.pack(pady=10)
        
        # Etiqueta de estado
        self.status_label = tk.Label(
            main_frame,
            text="Listo para procesar",
            font=("Arial", 9),
            fg="#555555"
        )
        self.status_label.pack(pady=(0, 10))
        
        # Área de registro
        log_label = tk.Label(
            main_frame,
            text="Registro de actividad:",
            font=("Arial", 10)
        )
        log_label.pack(anchor="w", pady=(20, 5))
        
        self.log_area = scrolledtext.ScrolledText(
            main_frame,
            height=8,
            font=("Arial", 9)
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.configure(state="disabled")
        
    def log_message(self, message):
        self.log_area.configure(state="normal")
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(END, f"[{timestamp}] {message}\n")
        self.log_area.see(END)
        self.log_area.configure(state="disabled")
        self.root.update_idletasks()
    
    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def load_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        
        if files:
            for file_path in files:
                if file_path not in self.selected_files:
                    self.selected_files.append(file_path)
                    self.files_listbox.insert(END, os.path.basename(file_path))
            self.log_message(f"Seleccionados {len(files)} archivos PDF")
    
    def remove_selected(self):
        selected_index = self.files_listbox.curselection()
        if selected_index:
            index = selected_index[0]
            removed_file = self.selected_files.pop(index)
            self.files_listbox.delete(index)
            self.log_message(f"Archivo eliminado: {os.path.basename(removed_file)}")
    
    def preprocess_image(self, image):
        """Mejora la imagen para OCR"""
        # Convertir a escala de grises
        img = image.convert('L')
        
        # Aumentar contraste
        img = img.point(lambda x: 0 if x < 128 else 255, '1')
        
        return img
    
    def pdf_to_text(self, pdf_path):
        """Convierte PDF a texto usando OCR"""
        try:
            # Crear directorio temporal
            with tempfile.TemporaryDirectory() as temp_dir:
                # Convertir PDF a imágenes
                self.update_status("Convirtiendo PDF a imágenes...")
                images = convert_from_path(
                    pdf_path,
                    dpi=300,  # Alta resolución para mejor OCR
                    output_folder=temp_dir,
                    fmt='jpeg',
                    thread_count=2
                )
                
                full_text = ""
                
                # Procesar cada imagen con OCR
                for i, image in enumerate(images):
                    self.update_status(f"Procesando página {i+1}/{len(images)}...")
                    
                    # Preprocesar imagen
                    processed_image = self.preprocess_image(image)
                    
                    # Usar OCR para extraer texto
                    text = pytesseract.image_to_string(
                        processed_image, 
                        lang='spa'  # Especificar idioma español
                    )
                    full_text += text + "\n\n"
                
                return full_text
        except Exception as e:
            self.log_message(f"Error en conversión OCR: {str(e)}")
            return ""
    
    def extract_pdf_info(self, pdf_path):
        try:
            # Extraer texto con OCR
            text = self.pdf_to_text(pdf_path)
            
            if not text:
                return None
            
            # Guardar texto extraído para depuración
            debug_file = os.path.join(self.results_folder, "debug_ocr_text.txt")
            with open(debug_file, "w", encoding="utf-8") as f:
                f.write(text)
            
            # Extraer número de serie (formato: 5-xxx/xxx)
            serial_number = "NO ENCONTRADO"
            
            # Patrones para número de serie
            patterns = [
                r'5-\d+/\d+',          # Formato estándar
                r'5\s*[-]\s*\d+\s*/\s*\d+',  # Con espacios
                r'5\s*[-]\s*\d+\s*[-]\s*\d+', # Con guiones
                r'5\s*[-/]\s*\d+\s*[-/]\s*\d+' # Variantes de separadores
            ]
            
            for pattern in patterns:
                serial_match = re.search(pattern, text)
                if serial_match:
                    # Limpiar espacios en el número encontrado
                    serial_number = re.sub(r'\s+', '', serial_match.group(0))
                    break
            
            # Extraer asunto
            asunto = "NO ENCONTRADO"
            
            # Patrones para asunto (mejorados)
            asunto_patterns = [
                r'Asunto\s*[:]\s*(.*?)(?=Cordial Saludo|Atento|Respetuoso|Estimado|$)',  # Hasta saludos o fin
                r'Asunto\s*[:]\s*(.*?)(?=\n\n|\n\s*\n|$)',    # Hasta doble salto de línea
                r'Asunto\s*[:]\s*(.*)'                        # Todo lo que sigue
            ]
            
            for pattern in asunto_patterns:
                asunto_match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
                if asunto_match:
                    asunto = asunto_match.group(1).strip()
                    # Limpiar espacios múltiples y saltos de línea
                    asunto = re.sub(r'\s+', ' ', asunto)
                    # Eliminar posibles firmas o sellos al final
                    asunto = re.sub(r'[\s\-_]+$', '', asunto)
                    break
            
            return {
                'Archivo PDF': os.path.basename(pdf_path),
                'Número de Serie': serial_number,
                'Asunto': asunto,
                'Texto Extraído': text  # Para depuración
            }
        except Exception as e:
            self.log_message(f"Error procesando {os.path.basename(pdf_path)}: {str(e)}")
            return None
    
    def open_excel_file(self):
        """Abre el archivo Excel generado con la aplicación predeterminada"""
        if not self.current_excel_file or not os.path.exists(self.current_excel_file):
            messagebox.showwarning("Advertencia", "No hay archivo Excel disponible para abrir")
            return
        
        try:
            if platform.system() == 'Darwin':       # macOS
                subprocess.call(('open', self.current_excel_file))
            elif platform.system() == 'Windows':    # Windows
                os.startfile(self.current_excel_file)
            else:                                   # linux variants
                subprocess.call(('xdg-open', self.current_excel_file))
                
            self.log_message(f"Archivo Excel abierto: {self.current_excel_file}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo Excel: {str(e)}")
            self.log_message(f"Error al abrir Excel: {str(e)}")
    
    def process_files(self):
        if not self.selected_files:
            messagebox.showwarning("Advertencia", "No se han seleccionado archivos PDF")
            return
        
        try:
            # Crear nombre de archivo basado en la fecha
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            self.current_excel_file = os.path.join(self.results_folder, f"oficios_{today}.xlsx")
            
            # Verificar si ya existe el archivo Excel
            if os.path.exists(self.current_excel_file):
                df = pd.read_excel(self.current_excel_file)
            else:
                df = pd.DataFrame(columns=['Archivo PDF', 'Número de Serie', 'Asunto'])
            
            # Configurar barra de progreso
            self.progress['value'] = 0
            self.progress['maximum'] = len(self.selected_files)
            
            # Procesar cada archivo
            success_count = 0
            for i, file_path in enumerate(self.selected_files):
                filename = os.path.basename(file_path)
                self.log_message(f"Procesando: {filename}...")
                self.update_status(f"Procesando: {filename}...")
                self.root.update_idletasks()
                
                info = self.extract_pdf_info(file_path)
                if info:
                    # Agregar al DataFrame, excluyendo el texto completo para el Excel
                    row = {
                        'Archivo PDF': info['Archivo PDF'],
                        'Número de Serie': info['Número de Serie'],
                        'Asunto': info['Asunto']
                    }
                    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
                    success_count += 1
                    self.log_message(f"  → N° Serie: {info['Número de Serie']}")
                    self.log_message(f"  → Asunto: {info['Asunto']}")
                else:
                    self.log_message("  → No se pudo extraer información")
                
                # Actualizar progreso
                self.progress['value'] = i + 1
                self.root.update_idletasks()
            
            # Guardar resultados
            df.to_excel(self.current_excel_file, index=False)
            
            # Mostrar resultados
            self.log_message(f"Procesamiento completado: {success_count}/{len(self.selected_files)} archivos procesados")
            self.log_message(f"Resultados guardados en: {self.current_excel_file}")
            self.update_status("Procesamiento completado")
            
            # Habilitar botón de abrir Excel
            self.open_excel_button.config(state=tk.NORMAL)
            
            # Reiniciar selección
            self.selected_files = []
            self.files_listbox.delete(0, END)
            self.progress['value'] = 0
            
            messagebox.showinfo(
                "Proceso completado",
                f"Se procesaron {success_count} de {len(self.selected_files)} archivos\n\n"
                f"Resultados guardados en:\n{self.current_excel_file}"
            )
            
        except Exception as e:
            self.update_status("Error durante el procesamiento")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
            self.log_message(f"Error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()