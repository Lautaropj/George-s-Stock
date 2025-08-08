
import pandas as pd
from datetime import datetime
import json

fecha_ahora = datetime.now()
hora_y_minutos = fecha_ahora.strftime("%H:%M")
hoy = fecha_ahora.strftime("%d/%m")

#El archivo siempre tiene el nombre 'Ventas 01.MM.AA.xlsm'
fecha_archivo = fecha_ahora.strftime(f"01.%m.%y") 

#Abrir el archivo de Excel
try:
    df = pd.read_excel(f'C:\\Users\\Usuario\\OneDrive\\Ventas {fecha_archivo}.xlsm', sheet_name='Stock', header=1)
except FileNotFoundError:
    print(f"El archivo 'Ventas {fecha_archivo}.xlsm' no se encuentra en la ruta especificada.")

# Convertir la columna 'FECHA' a tipo fecha y filtrar por la fecha actual
df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce').dt.date
df_filtrado = df[df['FECHA'] == fecha_ahora.date()]

df_filtrado = df_filtrado.map(lambda x:x.strip() if isinstance(x, str) else x)

#Abrir JSON necesarios
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
        alias_productos = json.load(f)
except FileNotFoundError:
    mostrar_error("El archivo 'alias_productos.json' no se encontró.")
    exit()
    

#lista que contendrá los productos con stock bajo
productos_bajos_totales = []

for index, row in df_filtrado.iterrows():
    producto = row['PRODUCTO']
    if pd.isna(producto) or pd.isna(row['STOCK']):
        continue
    stock_actual = row['STOCK']
    if not isinstance(stock_actual, (int, float)):
        continue
    if producto in stock_conditions:
        limite, ideal = stock_conditions[producto]
        if stock_actual <= limite:
            cantidad_necesaria = ideal - stock_actual
            productos_bajos_totales.append((producto, stock_actual, cantidad_necesaria))

def generar_reporte():
    if not productos_bajos_totales:
        print("No hay productos con stock bajo.")
        return
    
    mensaje = f"Reporte de productos con stock bajo - {hoy} a las {hora_y_minutos}\n\n"
    
    for producto, cantidad_actual, cantidad_necesaria in productos_bajos_totales:
        producto_original = producto
        
        if producto in alias_productos:
            producto = alias_productos[producto]

        if producto in productos_con_plus:
            mensaje += f"{producto}: {cantidad_actual} + {cantidad_necesaria}\n"
        else:
            mensaje += f"{producto}: {cantidad_actual}\n"

    with open('C:\\Users\\Usuario\\OneDrive\\Escritorio\\Stock generado2.txt', 'w', encoding="utf-8") as file:
        file.write(mensaje)

generar_reporte()



