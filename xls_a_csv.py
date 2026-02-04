import pandas as pd
import unidecode
import sys
import os
from glob import glob
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox


def limpiar_texto(valor):
    if isinstance(valor, str):
        valor = valor.strip()
        valor = unidecode.unidecode(valor)
    return valor

def limpiar_numero(valor):
    if isinstance(valor, str):
        valor = valor.replace(",", ".")
        try:
            valor = float(valor)
        except ValueError:
            pass
    return valor

def limpiar_fecha(valor):
    try:
        dt = pd.to_datetime(valor)
        if dt.time() != pd.Timestamp(0).time():
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        else:
            return dt.strftime("%Y-%m-%d")
    except:
        return valor


def procesar_excel(carpeta_entrada, carpeta_salida, log_widget):
    archivos = glob(os.path.join(carpeta_entrada, "*.xls"))
    if not archivos:
        log_widget.insert(tk.END, "[INFO] No se encontraron archivos .xls en la carpeta.\n")
        return
    
    for ruta_archivo in archivos:
        try:
            df = pd.read_excel(ruta_archivo, engine='xlrd')
            
            
            for columna in df.columns:
                df[columna] = df[columna].apply(limpiar_texto)
                df[columna] = df[columna].apply(limpiar_numero)
                df[columna] = df[columna].apply(limpiar_fecha)
            
            
            nombre_csv = os.path.splitext(os.path.basename(ruta_archivo))[0] + ".csv"
            df.to_csv(os.path.join(carpeta_salida, nombre_csv), index=False, encoding="utf-8-sig")
            log_widget.insert(tk.END, f"[OK] {nombre_csv} generado.\n")
            log_widget.see(tk.END)
            
        except Exception as e:
            log_widget.insert(tk.END, f"[ERROR] Fallo al procesar {ruta_archivo}: {e}\n")
            log_widget.see(tk.END)


def seleccionar_carpeta_entrada():
    carpeta = filedialog.askdirectory()
    if carpeta:
        entrada_var.set(carpeta)

def seleccionar_carpeta_salida():
    carpeta = filedialog.askdirectory()
    if carpeta:
        salida_var.set(carpeta)

def boton_procesar():
    carpeta_entrada = entrada_var.get()
    carpeta_salida = salida_var.get()
    
    if not carpeta_entrada or not carpeta_salida:
        messagebox.showwarning("Advertencia", "Seleccione las carpetas de entrada y salida.")
        return
    
    log_text.delete(1.0, tk.END)
    procesar_excel(carpeta_entrada, carpeta_salida, log_text)
    messagebox.showinfo("Finalizado", "Proceso completado.")


root = tk.Tk()

def ruta_recurso(ruta_relativa):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, ruta_relativa)
    return os.path.join(os.path.abspath("."), ruta_relativa)

root.iconbitmap(ruta_recurso("icono.ico"))

root.title("Procesador de XLS a CSV")


entrada_var = tk.StringVar()
salida_var = tk.StringVar()


frame = tk.Frame(root, padx=10, pady=10)
frame.pack()


tk.Label(frame, text="Carpeta de entrada:").grid(row=0, column=0, sticky="w")
tk.Entry(frame, textvariable=entrada_var, width=50).grid(row=0, column=1)
tk.Button(frame, text="Seleccionar", command=seleccionar_carpeta_entrada).grid(row=0, column=2, padx=5)


tk.Label(frame, text="Carpeta de salida:").grid(row=1, column=0, sticky="w")
tk.Entry(frame, textvariable=salida_var, width=50).grid(row=1, column=1)
tk.Button(frame, text="Seleccionar", command=seleccionar_carpeta_salida).grid(row=1, column=2, padx=5)


tk.Button(frame, text="Procesar XLS", command=boton_procesar, bg="lightgreen").grid(row=2, column=0, columnspan=3, pady=10)


log_text = scrolledtext.ScrolledText(frame, width=80, height=20)
log_text.grid(row=3, column=0, columnspan=3, pady=5)


root.mainloop()