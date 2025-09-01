import pandas as pd
from tkinter import messagebox
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
from copy import deepcopy
from tkinter import filedialog
import os



def extraer_datos(archivo, pdf):

    ruta_salida = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("Archivos PDF", "*.pdf")],
        title="Guardar PDF generado como...")

    try:
        # Leer Excel
        df = pd.read_excel(archivo, usecols=[0,1,2], header=None)
        df = df.dropna(how='all').reset_index(drop=True)

        # Última fila con datos en la segunda columna
        ultimo_indice = df[2].last_valid_index()
        df_filtrado = df.loc[:ultimo_indice]

        # Definir bloques de 10 registros
        num_registros = len(df_filtrado) // 20

        # Cargar la plantilla PDF
        plantilla = PdfReader(pdf)
        writer = PdfWriter()

        for i in range(num_registros):
            bloque = df_filtrado.iloc[i * 20:(i + 1) * 20]
            print(f"\nBloque {i + 1}:\n", bloque)

            # Definir tamaño de página 6x4 pulgadas
            ancho = 4 * 72
            alto = 6 * 72

            # Crear PDF temporal con datos
            temp_pdf = f"temp_{i + 1}.pdf"
            c = canvas.Canvas(temp_pdf, pagesize=(ancho, alto))
            c.setFont("Helvetica-BoldOblique", 11)

            y = alto - 30
            for _, row in bloque.iterrows():
                etiqueta = str(row[0])
                valor = str(row[2])

                #  parte 1 del pdf
                if(etiqueta=="fecha1"):
                    c.drawString(200,416, f" {valor}")
                if(etiqueta=="nombre1"):
                    c.drawString(67, 377, f" {valor}")
                if(etiqueta=="telefono1"):
                    c.drawString(70, 360, f" {valor}")
                if(etiqueta=="calle1"):
                    c.drawString(52, 345, f" {valor}")
                if(etiqueta=="colonia1"):
                    if len(valor)>37:
                        cadena = division_linea(valor)
                        c.drawString(65, 330, f" {cadena[0]}")
                        c.drawString(65, 320, f" {cadena[1]}")
                    else:
                        c.drawString(65, 330, f" {valor}")
                if(etiqueta=="ref1"):
                    if (len(valor)>37):
                        cadena = division_linea(valor)
                        c.drawString(78, 307, f" {cadena[0]}")
                        c.drawString(78, 297, f" {cadena[1]}")
                    else:
                        c.drawString(78, 307, f" {valor}")
                if(etiqueta=="costo1"):
                    c.drawString(72, 283, f"$ {valor}")
                if(etiqueta=="envio1"):
                    c.drawString(55, 267, f"$ {valor}")
                if(etiqueta=="total1"):
                    c.drawString(55, 253, f"$ {valor}")
                if(etiqueta=="nota1"):
                    if(len(valor)>37):
                        cadena = division_linea(valor)
                        c.drawString(50, 237, f" {cadena[0]}")
                        c.drawString(50, 227, f" {cadena[1]}")
                    else:
                        c.drawString(50, 237, f" {valor}")

                #  parte 2 del pdf
                if(etiqueta=="fecha2"):
                    c.drawString(200,196, f" {valor}")
                if(etiqueta=="nombre2"):
                    c.drawString(67, 154, f" {valor}")
                if(etiqueta=="telefono2"):
                    c.drawString(70, 139, f" {valor}")
                if(etiqueta=="calle2"):
                    c.drawString(52, 124, f" {valor}")
                if(etiqueta=="colonia2"):
                    if len(valor)>37:
                        cadena = division_linea(valor)
                        c.drawString(65, 109, f" {cadena[0]}")
                        c.drawString(65, 99, f" {cadena[1]}")
                    else:
                        c.drawString(65, 109, f" {valor}")
                if(etiqueta=="ref2"):
                    if (len(valor)>37):
                        cadena = division_linea(valor)
                        c.drawString(78, 86, f" {cadena[0]}")
                        c.drawString(78, 76, f" {cadena[1]}")
                    else:
                        c.drawString(78, 86, f" {valor}")
                if(etiqueta=="costo2"):
                    c.drawString(72, 62, f"$ {valor}")
                if(etiqueta=="envio2"):
                    c.drawString(55, 46, f"$ {valor}")
                if(etiqueta=="total2"):
                    c.drawString(55, 32, f"$ {valor}")
                if(etiqueta=="nota2"):
                    if(len(valor)>37):
                        cadena = division_linea(valor)
                        c.drawString(50, 16, f" {cadena[0]}")
                        c.drawString(50, 6, f" {cadena[1]}")
                    else:
                        c.drawString(50, 16, f" {valor}")

            c.save()

            # Fusionar con plantilla
            pagina_plantilla = deepcopy(plantilla.pages[0])
            overlay = PdfReader(temp_pdf).pages[0]
            pagina_plantilla.merge_page(overlay)
            writer.add_page(pagina_plantilla)

            os.remove(temp_pdf)

        # Guardar resultado final
        if ruta_salida:
            with open(ruta_salida, "wb") as salida:
                writer.write(salida)
                os.startfile(ruta_salida)

        messagebox.showinfo(
            "Éxito",
            f"Se generó el PDF con {num_registros} páginas:\n{ruta_salida}"
        )

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo Excel: {e}")


def division_linea(valor):
    corte = valor.rfind("",0,37)
    if corte == -1:
        corte = 37

    parte1 = valor[:corte].strip()
    parte2 = valor[corte:].strip()

    return [parte1, parte2]
