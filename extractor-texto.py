import os
import re
import datetime
import PyPDF2
import pandas as pd
from tkinter import Tk, filedialog, messagebox

def extract_pdf_info(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        first_page = reader.pages[0]
        text = first_page.extract_text()
        
        # Extraer número de serie (formato: 5-xxx/xxx)
        serial_match = re.search(r'5-\d+/\d+', text)
        serial_number = serial_match.group(0) if serial_match else "NO ENCONTRADO"
        
        # Extraer asunto (entre "Asunto:" y "Cordial Saludo")
        asunto_match = re.search(r'Asunto:(.*?)Cordial Saludo', text, re.DOTALL | re.IGNORECASE)
        asunto = asunto_match.group(1).strip() if asunto_match else "NO ENCONTRADO"
        
        # Limpiar texto del asunto
        asunto = ' '.join(asunto.split()).replace('\n', ' ')
        
        return {
            'Archivo PDF': os.path.basename(pdf_path),
            'Número de Serie': serial_number,
            'Asunto': asunto
        }

def process_files():
    root = Tk()
    root.withdraw()
    
    files = filedialog.askopenfilenames(
        title="Seleccionar PDFs",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    
    if not files:
        return
    
    today = datetime.datetime.now().strftime("%Y%m%d")
    excel_file = f"oficios_{today}.xlsx"
    
    try:
        # Verificar si el archivo Excel ya existe
        if os.path.exists(excel_file):
            df = pd.read_excel(excel_file)
        else:
            df = pd.DataFrame(columns=['Archivo PDF', 'Número de Serie', 'Asunto'])
        
        # Procesar cada archivo PDF
        for file_path in files:
            try:
                info = extract_pdf_info(file_path)
                df = pd.concat([df, pd.DataFrame([info])], ignore_index=True)
            except Exception as e:
                print(f"Error procesando {file_path}: {str(e)}")
        
        # Guardar en Excel
        df.to_excel(excel_file, index=False)
        messagebox.showinfo("Éxito", f"Se procesaron {len(files)} archivos\nGuardado en: {excel_file}")
    
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

if __name__ == "__main__":
    process_files()