import tkinter as Tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from datos import extraccionDatos
import openpyxl



def seleccionar_excel():
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if archivo:
        ruta_excel.set(archivo)  # actualizar el Label
    else:
        ruta_excel.set("")       # si canceló, limpiar

def seleccionar_pdf():
    archivo = filedialog.askopenfilename(
        title="Seleccionar PDF plantilla",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if archivo:
        ruta_pdf.set(archivo)
    else:
        ruta_pdf.set("")


def procesar():
    archivo_excel = ruta_excel.get()
    archivo_pdf = ruta_pdf.get()

    if not archivo_excel:
        messagebox.showwarning("Advertencia", "Primero selecciona un archivo Excel.")
        return
    if not archivo_pdf:
        messagebox.showwarning("Advertencia", "Primero selecciona un PDF plantilla.")
        return
    
    extraccionDatos.extraer_datos(archivo_excel,archivo_pdf)




ventana = Tk.Tk()
ventana.title("Impresiones Chooms GDL")

#poscionar ventana en el centro
ancho_ventana = 600
alto_ventana = 180
ancho_pantalla = ventana.winfo_screenwidth()
alto_pantalla = ventana.winfo_screenheight()
pos_x = (ancho_pantalla // 2) - (ancho_ventana // 2)
pos_y = (alto_pantalla // 2) - (alto_ventana // 2)
ventana.geometry(f"{ancho_ventana}x{alto_ventana}+{pos_x}+{pos_y}")
ventana.configure(bg="#353535")

#configuracion del icono
ventana.iconbitmap("impresiones/chooms.ico")
#ventana no expandible
ventana.resizable(False, False)

# Selección de Excel
frame_excel = ttk.Frame(ventana)
frame_excel.pack(pady=5)
btn_excel = ttk.Button(frame_excel, text="Seleccionar Excel", command=seleccionar_excel)
btn_excel.grid(row=0, column=0, padx=5)
ruta_excel = Tk.StringVar()
lbl_excel = ttk.Label(frame_excel, textvariable=ruta_excel, wraplength=400, anchor="w", justify="left")
lbl_excel.grid(row=0, column=1, padx=5)

# Selección de PDF plantilla
frame_pdf = ttk.Frame(ventana)
frame_pdf.pack(pady=5)
btn_pdf = ttk.Button(frame_pdf, text="Seleccionar PDF plantilla", command=seleccionar_pdf)
btn_pdf.grid(row=0, column=0, padx=5)
ruta_pdf = Tk.StringVar()
lbl_pdf = ttk.Label(frame_pdf, textvariable=ruta_pdf, wraplength=400, anchor="w", justify="left")
lbl_pdf.grid(row=0, column=1, padx=5)

# Botón procesar
btn_procesar = ttk.Button(ventana, text="Generar PDF", command=procesar)
btn_procesar.pack(pady=15)



ventana.mainloop()
