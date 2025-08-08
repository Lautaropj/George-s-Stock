import os
import pandas as pd
import json
from datetime import datetime
import time
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

def mostrar_error(mensaje):  # Función para mostrar mensaje de error
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", mensaje)
    root.quit()

fecha_ahora = datetime.now()  # Fecha y hora actual
minutos = fecha_ahora.strftime("%H:%M")
hoy = fecha_ahora.strftime('%d/%m')  
fecha_actual = datetime.today()
nombre_archivo = fecha_actual.strftime("01.%m.%y") #El archivo siempre tiene el nombre "Ventas 01.MM.YY.xslm"

try:
    df = pd.read_excel(f'C:\\Users\\Usuario\\OneDrive\\Ventas {nombre_archivo}.xlsm', header=1, sheet_name="Stock")  # Cargar el archivo Excel
except FileNotFoundError:
    mostrar_error(f"No se ha encontrado el archivo Excel con el nombre 'Ventas {nombre_archivo}.xlsm', o el nombre del mismo es incorrecto.\nEl mismo debe tener como nombre el primer día del mes.\nEjemplo válido: 'Ventas 01.04.25.xlsm'")
    exit()

#Filtrado y limpieza
df.columns = df.columns.str.strip()  # Eliminar espacios extras en los nombres de las columnas
df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')  # Convertir la columna 'FECHA' a datetime
df_filtrado = df[df['FECHA'].dt.date == fecha_ahora.date()]  # Filtrar por la fecha actual
df_filtrado = df_filtrado.applymap(lambda x: x.strip() if isinstance(x, str) else x)  # Limpiar espacios adicionales en los datos

try:
    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\stock_conditions.json', 'r', encoding="utf-8") as f:  # Cargar condiciones de stock
        stock_conditions = json.load(f)
except FileNotFoundError:
    mostrar_error("El archivo 'stock_conditions.json' no se encontró.")
    exit()

try:
    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\productos_con_plus.json', 'r', encoding="utf-8") as f:  # Cargar productos con plus
        data = json.load(f)
        productos_con_plus = data["productos_con_plus"]
except FileNotFoundError:
    mostrar_error("El archivo 'productos_con_plus.json' no se encontró.")
    exit()

try:
    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\alias_productos.json', 'r', encoding="utf-8") as f:  # Cargar alias de productos
        mapeo_productos = json.load(f)
except FileNotFoundError:
    mostrar_error("El archivo 'alias_productos.json' no se encontró.")
    exit()

productos_bajos_totales = []  # Lista para productos con stock bajo
productos_discontinuos = []  # Lista para productos discontinuos

# Verificar cada producto y su stock
for index, row in df_filtrado.iterrows():
    producto = row['PRODUCTO']  
    if pd.isna(producto) or pd.isna(row['STOCK']):  # Si producto o stock son NaN, continuar
        continue

    stock_actual = row['STOCK']
    if not isinstance(stock_actual, (int, float)):  # Verificar si el stock es numérico
        print(f"Producto '{producto}' tiene un stock no válido: {stock_actual}")
        continue

    if producto in stock_conditions:  # Si el producto tiene condición de stock
        limite, ideal = stock_conditions[producto]
        if stock_actual <= limite:  # Si el stock es menor o igual al límite
            cantidad_necesaria = ideal - stock_actual  # Calcular la cantidad necesaria
            productos_bajos_totales.append((producto, int(stock_actual), int(cantidad_necesaria)))  # Agregar a la lista de productos bajos
    else:
        productos_discontinuos.append(producto)  # Si no tiene condición, agregar a discontinuos

def abrir_condiciones_stock():  # Función para abrir el archivo de condiciones de stock
    archivo = 'C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\stock_conditions.json'
    if os.path.exists(archivo):
        os.startfile(archivo)  # Abrir el archivo
    else:
        messagebox.showerror("Error", "No se pudo encontrar el archivo 'stock_conditions.json'.")

def abrir_productos_con_plus():  # Función para abrir el archivo de productos con plus
    archivo = 'C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\productos_con_plus.json'
    if os.path.exists(archivo):
        os.startfile(archivo)  # Abrir el archivo
    else:
        messagebox.showerror("Error", "No se pudo encontrar el archivo 'productos_con_plus.json'.")

def generar_stock_bajo(ventana):  # Función para generar el reporte de productos con stock bajo
    if not productos_bajos_totales:
        messagebox.showerror("Error", "No hay productos con bajo stock.")
        return

    mensaje = f'Productos con stock bajo para el {hoy} a las {minutos}\n\n'
    for producto, stock_actual, cantidad_necesaria in productos_bajos_totales:  # Iterar por los productos bajos
        producto_original = producto

        if producto in productos_con_plus:
            mensaje += f'{producto}: {stock_actual} +{cantidad_necesaria}\n'  # Si tiene un plus, mostrar la cantidad necesaria
        else:
            mensaje += f'{producto}: {stock_actual}\n'

        if producto in mapeo_productos:  # Si hay un alias, usarlo
            producto = mapeo_productos[producto]

        if producto != producto_original:
            mensaje = mensaje.replace(producto_original, producto)

    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Stock generado.txt', 'w', encoding="utf-8") as file:  # Guardar el reporte
        file.write(mensaje)

    def salir_programa():  # Función para salir del programa
        ventana.quit()
        ventana.destroy()
    
    def abrir_stock_bajo_y_salir():  # Función para abrir el archivo generado y salir
        archivo = 'C:\\Users\\Usuario\\OneDrive\\Escritorio\\Stock generado.txt'
        if os.path.exists(archivo):
            os.startfile(archivo)
            ventana_info.quit() 
            ventana_info.destroy()
            salir_programa()
        else:
            messagebox.showerror("Error", "No se pudo encontrar el archivo 'Stock generado.txt'.")

    ventana_info = tk.Toplevel()  # Crear ventana secundaria
    ventana_info.geometry("470x120")
    ventana_info.iconbitmap('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\en-stock3.ico')

    # Calcular posición de la ventana
    pantalla_ancho = ventana_info.winfo_screenwidth()
    pantalla_alto = ventana_info.winfo_screenheight()
    ventana_ancho = 600
    ventana_alto = 120

    posicion_x = (pantalla_ancho // 2) - (ventana_ancho // 2)
    posicion_y = (pantalla_alto // 2) - (ventana_alto // 2)

    ventana_info.geometry(f"{ventana_ancho}x{ventana_alto}+{posicion_x}+{posicion_y}")

    label_info = tk.Label(ventana_info, text="Se ha generado el archivo en el escritorio.", font=("Arial", 12, "bold"), width=50)  # Mensaje de éxito
    label_info.pack(pady=20)

    btn_aceptar = tk.Button(ventana_info, text="Continuar en el programa", command=ventana_info.destroy, font=("Arial", 12, "bold"), bg="#7A651D", fg="white", relief="flat", bd=5, cursor="hand2")  # Botón para continuar
    btn_aceptar.pack(side="left", padx=20, pady=10)

    btn_salir = tk.Button(ventana_info, text="Salir del programa", command=lambda: [salir_programa(), ventana_info.quit()], font=("Arial", 12, "bold"), bg="#DC143C", fg="white", relief="flat", bd=5, cursor="hand2")  # Botón para salir
    btn_salir.pack(side="right", padx=20, pady=10)
   
    btn_abrir_stock_bajo = tk.Button(ventana_info, text="Abrir .txt y salir", command=abrir_stock_bajo_y_salir, font=("Arial", 12, "bold"), bg="#7A651D", fg="white", relief="flat", bd=5, cursor="hand2")  # Botón para abrir el archivo y salir
    btn_abrir_stock_bajo.pack(side="bottom", pady=10)
    
    ventana_info.mainloop()

def cerrar_ventana(ventana):  # Función para cerrar la ventana principal
    ventana.quit()
    ventana.destroy()

def mostrar_interfaz():  # Función principal para mostrar la interfaz gráfica
    ventana = tk.Tk()
    ventana.title("REPORTE DE STOCK BAJO")
    ventana.iconbitmap('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\en-stock3.ico')

    # Calcular tamaño y posición de la ventana
    screen_width = ventana.winfo_screenwidth()
    screen_height = ventana.winfo_screenheight()

    window_width = 450
    window_height = 670

    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 5   

    ventana.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    # Cargar imágenes
    img = Image.open("C:\\users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\lovelygeorge.png")
    img = img.resize((450, 200))
    img_tk = ImageTk.PhotoImage(img)

    label_imagen = tk.Label(ventana, image=img_tk)
    label_imagen.image = img_tk
    label_imagen.pack(pady=10)

    icono_advertencia = Image.open("C:\\users\\Usuario\\OneDrive\\Escritorio\\Nueva carpeta\\senal-de-advertencia.png")
    icono_advertencia = icono_advertencia.resize((16, 16))
    icono_advertencia_tk = ImageTk.PhotoImage(icono_advertencia)

    titulo_productos = tk.Label(ventana, text="PRODUCTOS SIN CONDICIÓN DE STOCK:", font=("Arial", 12, "bold"), anchor="w", fg="#9B7C18")
    titulo_productos.pack(pady=10, padx=20)

    frame_scroll = tk.Frame(ventana, height=300)  # Crear contenedor con scroll
    frame_scroll.pack(pady=10, padx=20, fill="x", expand=False)

    canvas = tk.Canvas(frame_scroll, height=250)
    scrollbar = tk.Scrollbar(frame_scroll, orient="vertical", command=canvas.yview)
    canvas.config(yscrollcommand=scrollbar.set)

    frame_productos = tk.Frame(canvas)
    frame_productos.pack(fill="x", expand=False)

    canvas.create_window((0, 0), window=frame_productos, anchor="nw")

    if productos_discontinuos:  # Mostrar productos sin condición de stock
        for producto in productos_discontinuos:
            frame_producto = tk.Frame(frame_productos)
            frame_producto.pack(pady=5, anchor="w")

            label_producto = tk.Label(frame_producto, text=producto, font=("Arial", 12), image=icono_advertencia_tk, compound="left")
            label_producto.image = icono_advertencia_tk
            label_producto.pack()

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    # Crear botones de acción
    frame_botones = tk.Frame(ventana)
    frame_botones.place(x=42, y=window_height - 100)
    
    btn_generar = tk.Button(frame_botones, text="Generar stock bajo", command=lambda: generar_stock_bajo(ventana), font=("Arial", 11, "bold"), bg="#7A651D", fg="white", relief="flat", bd=3, width=18, height=1, cursor="hand2")
    btn_condiciones = tk.Button(frame_botones, text="Condiciones de stock", command=abrir_condiciones_stock, font=("Arial", 11, "bold"), bg="#7A651D", fg="white", relief="flat", bd=3, width=18, height=1, cursor="hand2")
    btn_plus = tk.Button(frame_botones, text='Productos con "+"', command=abrir_productos_con_plus, font=("Arial", 11, "bold"), bg="#7A651D", fg="white", relief="flat", bd=3, width=18, height=1, cursor="hand2")
    btn_salir = tk.Button(frame_botones, text="Salir", command=lambda: cerrar_ventana(ventana), font=("Arial", 11, "bold"), bg="#DC143C", fg="white", relief="flat", bd=3, width=18, height=1, cursor="hand2")

    btn_generar.grid(row=0, column=0, padx=5, pady=5)     
    btn_condiciones.grid(row=0, column=1, padx=5, pady=5)  
    btn_plus.grid(row=1, column=0, padx=5, pady=5)        
    btn_salir.grid(row=1, column=1, padx=5, pady=5)        

    ventana.mainloop()

mostrar_interfaz()  # Ejecutar la interfaz principal
