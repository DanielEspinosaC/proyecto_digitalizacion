import pandas as pd
import unidecode
import os
import sys
from glob import glob
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

# -----------------------
# Funciones de limpieza
# -----------------------

def limpiar_texto(valor):
    if isinstance(valor, str):
        if opcion_espacios.get():
            valor = valor.strip()
        if opcion_tildes.get():
            valor = unidecode.unidecode(valor)
    return valor

def limpiar_numero(valor):
    if isinstance(valor, str) and opcion_comas.get():
        valor = valor.replace(",", ".")
        try:
            valor = float(valor)
        except ValueError:
            pass
    return valor

def limpiar_fecha(valor):
    if not opcion_fechas.get():
        return valor
    try:
        dt = pd.to_datetime(valor)
        if dt.time() != pd.Timestamp(0).time():
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        else:
            return dt.strftime("%Y-%m-%d")
    except:
        return valor

# -----------------------
# Función principal
# -----------------------

def procesar_excel(carpeta_entrada, carpeta_salida, log_widget):
    archivos = glob(os.path.join(carpeta_entrada, "*.xls"))
    if not archivos:
        log_widget.insert(tk.END, "[INFO] No se encontraron archivos .xls en la carpeta.\n")
        return
    
    for ruta_archivo in archivos:
        try:
            df = pd.read_excel(ruta_archivo, engine="xlrd")

            for columna in df.columns:
                df[columna] = df[columna].apply(limpiar_texto)
                df[columna] = df[columna].apply(limpiar_numero)
                df[columna] = df[columna].apply(limpiar_fecha)

            nombre_csv = os.path.splitext(os.path.basename(ruta_archivo))[0] + ".csv"
            ruta_salida = os.path.join(carpeta_salida, nombre_csv)
            df.to_csv(ruta_salida, index=False, encoding="utf-8-sig")

            log_widget.insert(tk.END, f"[OK] {nombre_csv} generado.\n")
            log_widget.see(tk.END)

        except Exception as e:
            log_widget.insert(tk.END, f"[ERROR] Fallo al procesar {ruta_archivo}: {e}\n")
            log_widget.see(tk.END)

# -----------------------
# Funciones GUI
# -----------------------

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

# -----------------------
# Interfaz Tkinter
# -----------------------

root = tk.Tk()
root.title("Normalizador de Excel a CSV")

# Icono (si lo usas)
# root.iconbitmap("icono.ico")

# Tema oscuro / claro
tema_oscuro = tk.BooleanVar(value=False)

# Listas para mantener referencias a widgets (necesarias para cambiar estilos dinámicamente)
widget_labels = []
widget_entries = []
widget_buttons = []
widget_checkbuttons = []

# Definición de temas
light_theme = {
    "bg": "#f0f0f0",
    "frame_bg": "#f0f0f0",
    "fg": "#000000",
    "entry_bg": "#ffffff",
    "button_bg": "lightgreen",
    "text_bg": "#ffffff",
    "text_fg": "#000000",
}

dark_theme = {
    "bg": "#2b2b2b",
    "frame_bg": "#2b2b2b",
    "fg": "#e6e6e6",
    "entry_bg": "#3c3f41",
    "button_bg": "#5a8f7b",
    "text_bg": "#262626",
    "text_fg": "#e6e6e6",
}

def aplicar_tema(oscuro: bool):
    theme = dark_theme if oscuro else light_theme
    # Root and frame
    root.configure(bg=theme["bg"])
    frame.configure(bg=theme["frame_bg"])

    # Labels
    for lbl in widget_labels:
        try:
            lbl.configure(bg=theme["frame_bg"], fg=theme["fg"])
        except:
            pass

    # Entries
    for ent in widget_entries:
        try:
            ent.configure(bg=theme["entry_bg"], fg=theme["fg"], insertbackground=theme["fg"])
        except:
            pass

    # Buttons
    for btn in widget_buttons:
        try:
            btn.configure(bg=theme["button_bg"], fg=theme["fg"], activebackground=theme["button_bg"])
        except:
            pass

    # Checkbuttons
    for cb in widget_checkbuttons:
        try:
            cb.configure(bg=theme["frame_bg"], fg=theme["fg"], selectcolor=theme["frame_bg"])
        except:
            pass

    # Log (scrolledtext)
    try:
        log_text.configure(bg=theme["text_bg"], fg=theme["text_fg"], insertbackground=theme["text_fg"])
    except:
        pass

def toggle_tema():
    aplicar_tema(tema_oscuro.get())

entrada_var = tk.StringVar()
salida_var = tk.StringVar()

opcion_tildes = tk.BooleanVar(value=True)
opcion_comas = tk.BooleanVar(value=True)
opcion_espacios = tk.BooleanVar(value=True)
opcion_fechas = tk.BooleanVar(value=True)

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

# Carpetas
lbl = tk.Label(frame, text="Carpeta de entrada:")
lbl.grid(row=0, column=0, sticky="w")
widget_labels.append(lbl)
ent = tk.Entry(frame, textvariable=entrada_var, width=50)
ent.grid(row=0, column=1)
widget_entries.append(ent)
btn = tk.Button(frame, text="Seleccionar", command=seleccionar_carpeta_entrada)
btn.grid(row=0, column=2, padx=5)
widget_buttons.append(btn)

lbl = tk.Label(frame, text="Carpeta de salida:")
lbl.grid(row=1, column=0, sticky="w")
widget_labels.append(lbl)
ent = tk.Entry(frame, textvariable=salida_var, width=50)
ent.grid(row=1, column=1)
widget_entries.append(ent)
btn = tk.Button(frame, text="Seleccionar", command=seleccionar_carpeta_salida)
btn.grid(row=1, column=2, padx=5)
widget_buttons.append(btn)

# Opciones

lbl = tk.Label(frame, text="Opciones de normalizacion:")
lbl.grid(row=2, column=0, sticky="w", pady=(10,0))
widget_labels.append(lbl)

cb = tk.Checkbutton(frame, text="Quitar tildes", variable=opcion_tildes)
cb.grid(row=3, column=0, sticky="w")
widget_checkbuttons.append(cb)
cb = tk.Checkbutton(frame, text="Cambiar comas por puntos", variable=opcion_comas)
cb.grid(row=3, column=1, sticky="w")
widget_checkbuttons.append(cb)
cb = tk.Checkbutton(frame, text="Quitar espacios extra", variable=opcion_espacios)
cb.grid(row=4, column=0, sticky="w")
widget_checkbuttons.append(cb)
cb = tk.Checkbutton(frame, text="Normalizar fechas", variable=opcion_fechas)
cb.grid(row=4, column=1, sticky="w")
widget_checkbuttons.append(cb)

# Selector de tema
cb_tema = tk.Checkbutton(frame, text="Tema oscuro", variable=tema_oscuro, command=toggle_tema)
cb_tema.grid(row=2, column=2, sticky="e")
widget_checkbuttons.append(cb_tema)

# Boton
btn_proc = tk.Button(frame, text="Procesar Excel", command=boton_procesar, bg="lightgreen")
btn_proc.grid(row=5, column=0, columnspan=3, pady=10)
widget_buttons.append(btn_proc)

# Log
log_text = scrolledtext.ScrolledText(frame, width=80, height=20)
log_text.grid(row=6, column=0, columnspan=3)

# Aplicar tema inicial
aplicar_tema(tema_oscuro.get())

root.mainloop()