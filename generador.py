import os
import pandas as pd
import json
from datetime import datetime
import time
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import logging

def mostrar_error(mensaje):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", mensaje)
    root.quit()
    
try:
    logging.basicConfig(
        filename='C:\\Users\\Usuario\\OneDrive\\Escritorio\\Stock\\files\\log\\log.log', 
        level=logging.INFO,            
        format='%(asctime)s - %(levelname)s - %(message)s',  
        datefmt='%d-%m-%y %H:%M'     
    )
except FileNotFoundError:
    mostrar_error("Error al encontrar la carpeta de logs")
    exit()


fecha_ahora = datetime.now()
minutos = fecha_ahora.strftime("%H:%M")
hoy = fecha_ahora.strftime('%d/%m')  
fecha_actual = datetime.today()
nombre_archivo = fecha_actual.strftime("01.%m.%y")

try:
    df = pd.read_excel(f'C:\\Users\\Usuario\\OneDrive\\Ventas {nombre_archivo}.xlsm', header=1, sheet_name="Stock")
except FileNotFoundError:
    mostrar_error(f"No se ha encontrado el archivo Excel con el nombre 'Ventas {nombre_archivo}.xlsm', o el nombre del mismo es incorrecto.\nEl mismo debe tener como nombre el primer día del mes.\nEjemplo válido: 'Ventas 01.04.25.xlsm'")
    exit()

df.columns = df.columns.str.strip()

df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')

df_filtrado = df[df['FECHA'].dt.date == fecha_ahora.date()] 

df_filtrado = df_filtrado.applymap(lambda x: x.strip() if isinstance(x, str) else x)# Cargar los archivos JSON necesarios
try:
    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\stock_conditions.json', 'r', encoding="utf-8") as f:
        stock_conditions = json.load(f)
except FileNotFoundError:
    mostrar_error("El archivo 'stock_conditions.json' no se encontró.")
    exit()

try:
    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\productos_con_plus.json', 'r', encoding="utf-8") as f:
        data = json.load(f)
        productos_con_plus = data["productos_con_plus"]
except FileNotFoundError:
    mostrar_error("El archivo 'productos_con_plus.json' no se encontró.")
    exit()

try:
    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\alias_productos.json', 'r', encoding="utf-8") as f:
        mapeo_productos = json.load(f)
except FileNotFoundError:
    mostrar_error("El archivo 'alias_productos.json' no se encontró.")
    exit()

productos_bajos_totales = []
productos_discontinuos = []

for index, row in df_filtrado.iterrows():
    producto = row['PRODUCTO']  # Accede correctamente a la columna de producto
    if pd.isna(producto) or pd.isna(row['STOCK']):
        continue

    stock_actual = row['STOCK']
    if not isinstance(stock_actual, (int, float)):
        print(f"Producto '{producto}' tiene un stock no válido: {stock_actual}")
        continue

    if producto in stock_conditions:
        limite, ideal = stock_conditions[producto]
        if stock_actual <= limite:
            cantidad_necesaria = ideal - stock_actual
            productos_bajos_totales.append((producto, int(stock_actual), int(cantidad_necesaria)))
    else:
        productos_discontinuos.append(producto)

# Funciones relacionadas con la interfaz gráfica y demás operaciones
def abrir_condiciones_stock():
    archivo = 'C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\stock_conditions.json'
    if os.path.exists(archivo):
        os.startfile(archivo) 
    else:
        messagebox.showerror("Error", "No se pudo encontrar el archivo 'stock_conditions.json'.")

def abrir_productos_con_plus():
    archivo = 'C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\productos_con_plus.json'
    if os.path.exists(archivo):
        os.startfile(archivo) 
    else:
        messagebox.showerror("Error", "No se pudo encontrar el archivo 'productos_con_plus.json'.")
        
def generar_stock_bajo(ventana):
    if not productos_bajos_totales:
        messagebox.showerror("Error", "No hay productos con bajo stock.")
        return

    mensaje = f'Productos con stock bajo para el {hoy} a las {minutos}\n\n'
    for producto, stock_actual, cantidad_necesaria in productos_bajos_totales:
        producto_original = producto

        if producto in productos_con_plus:
            mensaje += f'{producto}: {stock_actual} +{cantidad_necesaria}\n'
        else:
            mensaje += f'{producto}: {stock_actual}\n'

        if producto in mapeo_productos:
            producto = mapeo_productos[producto]

        if producto != producto_original:
            mensaje = mensaje.replace(producto_original, producto)

    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Stock generado.txt', 'w', encoding="utf-8") as file:
        file.write(mensaje)
    
    logging.info(f"Se ha generado un reporte de stock")

    def salir_programa():
        ventana.quit()
        ventana.destroy()
    
    def abrir_stock_bajo_y_salir():
        archivo = 'C:\\Users\\Usuario\\OneDrive\\Escritorio\\Stock generado.txt'
        if os.path.exists(archivo):
            os.startfile(archivo)  # Abrir archivo de stock bajo generado
            ventana_info.quit()  # Cerrar la ventana de éxito
            ventana_info.destroy()
            salir_programa()
        else:
            messagebox.showerror("Error", "No se pudo encontrar el archivo 'Stock generado.txt'.")

    ventana_info = tk.Toplevel()
    ventana_info.geometry("470x120")
    ventana_info.iconbitmap('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\en-stock.ico')

    pantalla_ancho = ventana_info.winfo_screenwidth()
    pantalla_alto = ventana_info.winfo_screenheight()
    ventana_ancho = 600
    ventana_alto = 120

    posicion_x = (pantalla_ancho // 2) - (ventana_ancho // 2)
    posicion_y = (pantalla_alto // 2) - (ventana_alto // 2)

    ventana_info.geometry(f"{ventana_ancho}x{ventana_alto}+{posicion_x}+{posicion_y}")

    label_info = tk.Label(ventana_info, text="Se ha generado el archivo en el escritorio.", font=("Arial", 12, "bold"), width=50)
    label_info.pack(pady=20)

    btn_aceptar = tk.Button(ventana_info, text="Continuar en el programa", command=ventana_info.destroy, font=("Arial", 12, "bold"), bg="#2E8B57", fg="white", relief="flat", bd=5, cursor="hand2")
    btn_aceptar.pack(side="left", padx=20, pady=10)

    btn_salir = tk.Button(ventana_info, text="Salir del programa", command=lambda: [salir_programa(), ventana_info.quit()], font=("Arial", 12, "bold"), bg="#DC143C", fg="white", relief="flat", bd=5, cursor="hand2")
    btn_salir.pack(side="right", padx=20, pady=10)
   
    btn_abrir_stock_bajo = tk.Button(ventana_info, text="Abrir .txt y salir", command=abrir_stock_bajo_y_salir, font=("Arial", 12, "bold"), bg="#145cad", fg="white", relief="flat", bd=5, cursor="hand2")
    btn_abrir_stock_bajo.pack(side="bottom", pady=10)
    
    ventana_info.mainloop()

def cerrar_ventana(ventana):
    ventana.quit()
    ventana.destroy()

def mostrar_interfaz():
    ventana = tk.Tk()
    ventana.title("REPORTE DE STOCK BAJO")
    ventana.iconbitmap('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\en-stock.ico')

    screen_width = ventana.winfo_screenwidth()
    screen_height = ventana.winfo_screenheight()

    window_width = 450
    window_height = 670

    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 5   

    ventana.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    img = Image.open("C:\\users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\lovelygeorge.png")
    img = img.resize((450, 200))
    img_tk = ImageTk.PhotoImage(img)

    label_imagen = tk.Label(ventana, image=img_tk)
    label_imagen.image = img_tk
    label_imagen.pack(pady=10)

    icono_advertencia = Image.open("C:\\users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\senal-de-advertencia.png")
    icono_advertencia = icono_advertencia.resize((16, 16))
    icono_advertencia_tk = ImageTk.PhotoImage(icono_advertencia)

    titulo_productos = tk.Label(ventana, text="PRODUCTOS SIN CONDICIÓN DE STOCK:", font=("Arial", 12, "bold"), anchor="w", fg="#DC143C")
    titulo_productos.pack(pady=10, padx=20)

    frame_scroll = tk.Frame(ventana, height=300)
    frame_scroll.pack(pady=10, padx=20, fill="x", expand=False)

    canvas = tk.Canvas(frame_scroll, height=250)
    scrollbar = tk.Scrollbar(frame_scroll, orient="vertical", command=canvas.yview)
    canvas.config(yscrollcommand=scrollbar.set)

    frame_productos = tk.Frame(canvas)
    frame_productos.pack(fill="x", expand=False)

    canvas.create_window((0, 0), window=frame_productos, anchor="nw")

    if productos_discontinuos:
        for producto in productos_discontinuos:
            frame_producto = tk.Frame(frame_productos)
            frame_producto.pack(pady=5, anchor="w")

            label_producto = tk.Label(frame_producto, text=producto, font=("Arial", 12), image=icono_advertencia_tk, compound="left")
            label_producto.image = icono_advertencia_tk
            label_producto.pack()

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    frame_productos.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    frame_botones = tk.Frame(ventana)
    frame_botones.place(x=42, y=window_height - 100)
    
    btn_generar = tk.Button(frame_botones, text="Generar stock bajo", command=lambda: generar_stock_bajo(ventana), font=("Arial", 11, "bold"), bg="#2E8B57", fg="white", relief="flat", bd=3, width=18, height=1, cursor="hand2")
    btn_condiciones = tk.Button(frame_botones, text="Condiciones de stock", command=abrir_condiciones_stock, font=("Arial", 11, "bold"), bg="#145cad", fg="white", relief="flat", bd=3, width=18, height=1, cursor="hand2")
    btn_plus = tk.Button(frame_botones, text='Productos con "+"', command=abrir_productos_con_plus, font=("Arial", 11, "bold"), bg="#145cad", fg="white", relief="flat", bd=3, width=18, height=1, cursor="hand2")
    btn_salir = tk.Button(frame_botones, text="Salir", command=lambda: cerrar_ventana(ventana), font=("Arial", 11, "bold"), bg="#DC143C", fg="white", relief="flat", bd=3, width=18, height=1, cursor="hand2")

    btn_generar.grid(row=0, column=0, padx=5, pady=5)     
    btn_condiciones.grid(row=0, column=1, padx=5, pady=5)  
    btn_plus.grid(row=1, column=0, padx=5, pady=5)        
    btn_salir.grid(row=1, column=1, padx=5, pady=5)        

    ventana.mainloop()

mostrar_interfaz()
