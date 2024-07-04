import tkinter as tk
from tkinter import ttk
import openpyxl

def agregarDatos():
    nombres = entradaNombre.get()
    edad = int(spinboxEdad.get())
    suscripcion = estadoCombobox.get()
    empleado = "Empleado" if empleadoVar.get() else "Desempleado"

    # Insertar en excel
    archivoExcel = 'data/TestAPP.xlsx'
    workbook = openpyxl.load_workbook(archivoExcel)
    hoja = workbook.active
    
    valores = [nombres, edad, suscripcion, empleado]
    hoja.append(valores)
    workbook.save(archivoExcel)
    
    # Insertar en la tabla
    tabla.insert('', tk.END, values=valores)

def cambiarTema():
    """
    Cambia el tema de la aplicación entre 'forest-light' y 'forest-dark'
    basado en el estado del botón de cambio de tema.
    """
    if switchColorTema.instate(['selected']):
        estilo.theme_use('forest-light')
    else:
        estilo.theme_use('forest-dark')

def cargarDatos():
    """
    Carga los datos desde un archivo Excel y los imprime en la consola.
    """
    archivoExcel = 'data/TestAPP.xlsx'
    workbook = openpyxl.load_workbook(archivoExcel)
    hoja = workbook.active
    
    listaValores = list(hoja.values)
    
    for nombreColumna in listaValores[0]:
        tabla.heading(nombreColumna, text=nombreColumna)
    
    for valor in listaValores[1:]:
        tabla.insert('', tk.END, values=valor)

# Configuración inicial de la aplicación
aplicacion = tk.Tk()
aplicacion.resizable(False, False)
aplicacion.title("GestorPy")

# Configuración del estilo
estilo = ttk.Style(aplicacion)
aplicacion.call('source', 'forest-dark.tcl')
aplicacion.call('source', 'forest-light.tcl')

# Aplicar el tema 'forest-dark' inicialmente
estilo.theme_use('forest-dark')

# Variables
listaEstados = ["Suscrito", "No suscrito", "Otro"]

# UI
framePrincipal = ttk.Frame(aplicacion)
framePrincipal.pack()

frameEntradas = ttk.LabelFrame(framePrincipal, text="Gestor de Contenido")
frameEntradas.grid(row=0, column=0, padx=20, pady=10)

entradaNombre = ttk.Entry(frameEntradas)
entradaNombre.insert(0, "Nombre")
entradaNombre.bind('<FocusIn>', lambda e: entradaNombre.delete(0, tk.END))
entradaNombre.grid(row=0, column=0, sticky="ew", padx=(5, 5))

spinboxEdad = ttk.Spinbox(frameEntradas, from_=18, to=100)
spinboxEdad.insert(0, "Edad")
spinboxEdad.grid(row=1, column=0, sticky='ew', padx=5, pady=5)

estadoCombobox = ttk.Combobox(frameEntradas, values=listaEstados)
estadoCombobox.current(0)
estadoCombobox.grid(row=2, column=0, sticky='ew', padx=5, pady=5)

empleadoVar = tk.BooleanVar()
empleadoCheckbutton = ttk.Checkbutton(frameEntradas, text="Empleado", variable=empleadoVar)
empleadoCheckbutton.grid(row=3, column=0, sticky='nsew', padx=5, pady=5)

botonAgregar = ttk.Button(frameEntradas, text="Agregar", command=agregarDatos)
botonAgregar.grid(row=4, column=0, sticky='nsew', padx=5, pady=5)

lineaSeparadora = ttk.Separator(frameEntradas)
lineaSeparadora.grid(row=5, column=0, padx=(20, 10), pady=10, sticky='ew')

switchColorTema = ttk.Checkbutton(frameEntradas, text="Tema Blanco", style='Switch', command=cambiarTema)
switchColorTema.grid(row=6, column=0, padx=5, pady=5, sticky='nsew')

vistaContenidoTabla = ttk.Frame(framePrincipal)
vistaContenidoTabla.grid(row=0, column=1, pady=10)

# ScrollBar
scrollbar = ttk.Scrollbar(vistaContenidoTabla)
scrollbar.pack(side='right', fill='y')

# Configuración Columnas
columnas = ('Nombres', 'Edad', 'Suscripción', 'Empleado')
tabla = ttk.Treeview(vistaContenidoTabla, show='headings', columns=columnas, height=13, yscrollcommand=scrollbar.set)

# Configuración del tamaño de las columnas
tabla.column("Nombres", width=100)
tabla.column("Edad", width=50)
tabla.column("Suscripción", width=100)
tabla.column("Empleado", width=100)

tabla.pack()

# Cargar Datos
cargarDatos()

# Ejecución del bucle principal de la aplicación
aplicacion.mainloop()
