import os
import re
from tkinter import filedialog, messagebox

import pdfplumber
import pandas as pd
import openpyxl
import tkinter as tk


def extract_text_from_pdf(file_path):
    """
    This function opens the PDF and stores all the text from each page in a list.
    ----------------------
    Args:
        file_path: A string representing the path where the PDF file is located.
    Returns:
        A list containing the text from each page.
    """
    texts = []

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text())

    return texts


def extract_information(text):
    # Define regex patterns for extracting information
    punto_venta_pattern = r"NOTA DE CREDITO\s?N? (\d+) - (\d+)"
    numero_factura_pattern = r"NOTA DE CREDITO\s?N? (\d+) - (\d+)"
    fecha_emision_pattern = r"Fecha emisión: (\d{2}/\d{2}/\d{4})"
    importe_neto_pattern = r"Importe Neto ([\d.,]+)"
    iva_pattern = r"IVA (\d{2},\d{2})% ([\d.,]+)"
    impuesto_interno_pattern = r"IMPUESTO INTERNO (\d{2},\d{2})% ([\d.,]+)"
    extraer_domicilio_pattern = r'Domicilio:\s*(.*?)(?:\s+Ing\. Brutos N°|$)'
    extraer_codigo_postal_pattern = r'CP.\s*(.*?)(?:\s+Pedido Interno N°:|$)'
    nombre_empresa_pattern = r'Cliente código: (\d+)'

    # Extract matches using regex
    punto_venta_match = re.findall(punto_venta_pattern, text)
    numero_factura_match = re.findall(numero_factura_pattern, text)
    fecha_emision_match = re.findall(fecha_emision_pattern, text)
    importe_neto_match = re.findall(importe_neto_pattern, text)
    iva_match = re.findall(iva_pattern, text)
    impuesto_interno_match = re.findall(impuesto_interno_pattern, text)
    extraer_domicilio_match = re.findall(extraer_domicilio_pattern, text)
    extraer_codigo_postal_match = re.findall(extraer_codigo_postal_pattern, text)
    nombre_empresa_match = re.search(nombre_empresa_pattern, text)

    # Extract specific information from matches or use default values
    punto_venta = punto_venta_match[0][0] if punto_venta_match else "0"
    numero_factura = numero_factura_match[0][1] if numero_factura_match else "0"
    fecha_emision = fecha_emision_match[0] if fecha_emision_match else None
    importe_neto = importe_neto_match[0] if importe_neto_match else "0"
    alicuota, monto = iva_match[0] if iva_match else ("0", "0")
    alicuota_imp_interno, monto_imp_interno = impuesto_interno_match[0] if impuesto_interno_match else ("0", "0")
    domicilio = extraer_domicilio_match[0] if extraer_domicilio_match else "0"
    codigo_postal = extraer_codigo_postal_match[0] if extraer_codigo_postal_match else "0"
    nombre_empresa = nombre_empresa_match.group(1)

    # Return the extracted information as a dictionary
    return {
        "Numero de Factura": f"{punto_venta}-{numero_factura}",
        "Fecha de Emision": fecha_emision,
        "Nro Cliente": nombre_empresa,
        "Domicilio": f'{domicilio}, {codigo_postal}',
        "Importe Neto": importe_neto,
        "Monto de IVA": monto,
        "Monto de Imp Interno": monto_imp_interno
    }


def format_number(number):
    number = str(number).replace('.', '').replace(',', '.')
    return float(number)


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        archivo_label.configure(text=f"Archivo seleccionado: {file_path}", fg="green")
    else:
        archivo_label.configure(text="No se ha seleccionado ningún archivo", fg="red")

    return file_path


def start_extraction():
    file_path = select_file()
    text_pages = extract_text_from_pdf(file_path)

    invoices = [extract_information(text) for text in text_pages]

    # Creamos un DataFrame con la info extraida
    df_invoices = pd.DataFrame(invoices)

    # Apply the function to the desired columns
    df_invoices['Importe Neto'] = df_invoices['Importe Neto'].apply(format_number)
    df_invoices['Monto de IVA'] = df_invoices['Monto de IVA'].apply(format_number)
    df_invoices['Monto de Imp Interno'] = df_invoices['Monto de Imp Interno'].apply(format_number)

    # Exportar DataFrame a excel
    directory = os.path.dirname(file_path)  # Get the directory of the selected file
    output_file = os.path.join(directory, "invoice_to_excel.xlsx")
    df_invoices.to_excel(output_file, sheet_name="Facturas")
    messagebox.showinfo("Extraction Complete", "Extraction process finished.\n"
                                               f"The Excel file is saved at:\n{output_file}")


# Tkinter UI

# Create the Tkinter application window
window = tk.Tk()

# Set the window title
window.title("PDF Information Extraction")

window.geometry("500x250")

# Título de la ventana principal
titulo_label = tk.Label(window, text="Extraccion datos PDF",
                        font=("Arial", 18, "bold"))

titulo_label.pack(pady=10)

# Texto explicativo
explicacion_text = "1) Presione el botón 'Iniciar Proceso'\n" \
                   "2) Seleccione el archivo PDF.\n"

# Create a label for the explanatory text
explicacion_label = tk.Label(window, text=explicacion_text, font=("Arial", 9), justify="left")
explicacion_label.pack(pady=10)

# Create the "Start Extraction" button
extract_button = tk.Button(window, text="Iniciar Proceso", command=start_extraction)
extract_button.pack()

# Create a label to display the selected file
archivo_label = tk.Label(window, text="No se ha seleccionado ningún archivo", fg="red")
archivo_label.pack(pady=10)

# Footer
footer_label = tk.Label(window, text="Developed by Gian Franco Lorenzo")
footer_label.pack(side=tk.BOTTOM, pady=10)

# Start the Tkinter event loop
window.mainloop()
