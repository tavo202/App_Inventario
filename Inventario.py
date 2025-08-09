import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
from PIL import Image, ImageTk  # Necesita Pillow
import sys

# Variables globales
df = None
archivo_seleccionado = ""
checkbox_vars = []
ultima_carpeta = os.getcwd()
encabezado_img = None  # mantener referencia global



def resource_path(relative_path):
    """ Devuelve la ruta absoluta del recurso, compatible con ejecutable """
    try:
        base_path = sys._MEIPASS  # Carpeta temporal creada por PyInstaller
    except Exception:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)

def cargar_encabezado():
    global encabezado_img
    try:
        ruta_imagen = resource_path("fondo.jpeg")
        if not os.path.exists(ruta_imagen):
            messagebox.showerror("Error", f"No se encontró la imagen: {ruta_imagen}")
            return None
        imagen = Image.open(ruta_imagen)
        imagen = imagen.resize((300, 80), Image.Resampling.LANCZOS)
        encabezado_img = ImageTk.PhotoImage(imagen)
        return encabezado_img
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar la imagen de encabezado:\n{e}")
        return None


        

# Función para seleccionar archivo CSV
def seleccionar_archivo():
    global df, archivo_seleccionado, checkbox_vars, ultima_carpeta
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo CSV",
        filetypes=[("Archivos CSV", "*.csv")],
        initialdir=ultima_carpeta
    )
    if archivo:
        ultima_carpeta = os.path.dirname(archivo)
        archivo_seleccionado = archivo
        etiqueta_archivo.config(text=f"Archivo: {os.path.basename(archivo)}")
        try:
            df = pd.read_csv(archivo, encoding="utf-8", sep=",")
            if "Categoria" not in df.columns:
                messagebox.showerror("Error", "El archivo no contiene la columna 'Categoria'.")
                return

            for widget in frame_categorias.winfo_children():
                widget.destroy()
            checkbox_vars.clear()

            categorias = sorted(df["Categoria"].dropna().unique())

            for cat in categorias:
                var = tk.BooleanVar()
                chk = tk.Checkbutton(frame_categorias, text=cat, variable=var, anchor="w", justify="left")
                chk.pack(fill="x", padx=5, pady=2)
                checkbox_vars.append((cat, var))

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

# Función para tomar la selección y exportar Excel
def tomar_seleccion():
    global df, ultima_carpeta
    if df is None:
        messagebox.showwarning("Advertencia", "Primero selecciona un archivo CSV.")
        return

    categorias_seleccionadas = [cat for cat, var in checkbox_vars if var.get()]
    if not categorias_seleccionadas:
        messagebox.showwarning("Advertencia", "Selecciona al menos una categoría.")
        return

    df_filtrado = df[df["Categoria"].isin(categorias_seleccionadas)]

    columnas_exportar = ["REF", "Nombre", "Categoria", "En inventario [Arcaico café Bar]"]
    columnas_existentes = [col for col in columnas_exportar if col in df_filtrado.columns]
    df_exportar = df_filtrado[columnas_existentes]

    fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M")
    nombre_sugerido = f"Inventario_{fecha_hora}.xlsx"

    ruta_salida = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivo Excel", "*.xlsx")],
        title="Guardar archivo filtrado",
        initialdir=ultima_carpeta,
        initialfile=nombre_sugerido
    )
    if ruta_salida:
        try:
            df_exportar.to_excel(ruta_salida, index=False)
            messagebox.showinfo("Éxito", f"Archivo exportado en:\n{ruta_salida}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar:\n{e}")

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Control de inventario Arcaico")
ventana.geometry("500x600")

# Frame superior con imagen centrada
frame_encabezado = tk.Frame(ventana)
frame_encabezado.pack(pady=10)

img = cargar_encabezado()
if img:
    encabezado_label = tk.Label(frame_encabezado, image=img)
    encabezado_label.pack()

# Botón para seleccionar archivo
btn_seleccionar = tk.Button(ventana, text="Seleccionar archivo CSV", command=seleccionar_archivo)
btn_seleccionar.pack(pady=10)

# Etiqueta para mostrar archivo seleccionado
etiqueta_archivo = tk.Label(ventana, text="Ningún archivo seleccionado")
etiqueta_archivo.pack()

# Frame con scroll para categorías
frame_scroll = tk.Frame(ventana)
frame_scroll.pack(pady=10, fill="both", expand=True)

canvas = tk.Canvas(frame_scroll, highlightthickness=0)
scrollbar = tk.Scrollbar(frame_scroll, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

frame_categorias = scrollable_frame

# Botón para exportar
btn_tomar_seleccion = tk.Button(ventana, text="Tomar selección y exportar", command=tomar_seleccion)
btn_tomar_seleccion.pack(pady=10)

ventana.mainloop()
