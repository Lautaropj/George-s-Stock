# George-s-Stock
Este script automatiza la generación de un reporte de productos con bajo stock en base a un archivo de Excel de ventas. Utiliza una interfaz gráfica para mostrar advertencias y permitir acciones rápidas, como abrir archivos JSON de configuración o exportar el resultado a un archivo `.txt`.

## Funcionalidades principales

- Carga y analiza un archivo Excel mensual con los datos del stock.
- Verifica los niveles de stock de productos según condiciones predefinidas.
- Informa los productos cuyo stock está por debajo del mínimo establecido.
- Permite abrir archivos de configuración directamente desde la interfaz.
- Genera un archivo `.txt` con los productos bajos y lo guarda en el escritorio.
- Utiliza interfaz gráfica (`Tkinter`) para facilitar la interacción con el usuario.

## Requisitos

- Python 3.9 o superior.
- Librerías:
  - `pandas`
  - `openpyxl`
  - `tkinter` (viene con Python)
  - `Pillow`
  - `json`
  - `os`
  - `datetime`
  - `logging`
