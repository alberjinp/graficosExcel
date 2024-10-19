#-*- coding: Latin1 -*-

import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, simpledialog
import xlsxwriter
import matplotlib.patches as mpatches

# Desactivar el modo interactivo de Matplotlib
plt.ioff()

# Función para generar una gráfica en Python
def generar_grafica_excel(ruta_excel, hoja_nombre):
    # Cargar el archivo Excel y la hoja específica
    import openpyxl
    wb = openpyxl.load_workbook(ruta_excel)
    hoja = wb[hoja_nombre]
    
    df = pd.read_excel(ruta_excel, sheet_name=hoja_nombre, header=None, usecols="A:F", engine='openpyxl')

    # Obtener los encabezados de la fila 5 (Basal, Clínica, Funcionalidad, etc.)
    encabezados = df.iloc[4, 1:].tolist()

    # Obtener los nombres de las funciones desde A6 hasta A31
    funciones = df.iloc[5:31, 0].astype(str).tolist()

    # Obtener los valores de las series desde B6:F31
    datos = df.iloc[5:31, 1:].values

    # Solicitar el título del gráfico
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de Tkinter
    titulo_grafico = simpledialog.askstring("Título del Gráfico", "¿Cómo quieres llamar al gráfico?")

    # Solicitar el tipo de gráfico
    tipo_grafico = simpledialog.askstring("Tipo de Gráfico", "¿Qué tipo de gráfico deseas? (linea, barra, dispersión)").lower()

    # Crear la gráfica con Matplotlib
    colores_lineas = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
    fig, ax = plt.subplots(figsize=(30, 10))

    # Seleccionar el tipo de gráfico
    if tipo_grafico == 'barra':
        width = 0.15  # Ancho de cada barra
        separation = 0.05  # Separación entre grupos de barras
        x = range(len(funciones))
        for idx, columna in enumerate(encabezados):
            ax.bar([pos + idx * (width + separation) for pos in x], datos[:, idx], width=width, color=colores_lineas[idx % len(colores_lineas)], label=columna)
    elif tipo_grafico == 'dispersión':
        for idx, columna in enumerate(encabezados):
            ax.scatter(funciones, datos[:, idx], color=colores_lineas[idx % len(colores_lineas)], label=columna)
    else:
        for idx, columna in enumerate(encabezados):
            ax.plot(funciones, datos[:, idx], marker='o', color=colores_lineas[idx % len(colores_lineas)], markerfacecolor='white', markeredgecolor=colores_lineas[idx % len(colores_lineas)], linewidth=3.5, label=columna)  # Ensanchar el gráfico manteniendo la altura
    # Obtener los valores y colores desde B34:D39
    colores_referencia = {}
    for fila in range(34, 40):
        valor = hoja[f'B{fila}'].value
        color_hex = hoja[f'D{fila}'].value
        if not color_hex:
            color_hex = '#000000'
        print(f"Fila: {fila}, Valor: {valor}, Color (procesado): {color_hex}")
        
        
        colores_referencia[valor] = color_hex

    # Rellenar el fondo del gráfico en base a los colores de los valores
    for valor, color in colores_referencia.items():
        ax.axhspan(valor, valor + 1, facecolor=color, alpha=0.3)

    # Cambiar el color del texto de los rótulos del eje X en base a los valores de Basal
    valores_basal = df.iloc[5:31, 1].tolist()
    colores_rotulos = [colores_referencia.get(valor, '#000000') if pd.notna(valor) and isinstance(valor, (int, float)) else '#000000' for valor in valores_basal]
    plt.xticks(ticks=range(len(funciones)), labels=[funcion.replace(' ', '\n') if isinstance(funcion, str) and ' ' in funcion else (funcion if funcion is not None else '') for funcion in funciones], fontsize=8, ha='center')
    for tick_label, color in zip(ax.get_xticklabels(), colores_rotulos):
        tick_label.set_color(color)

    plt.xlabel('Funciones', fontsize=12)
    plt.ylabel('Valores (0 a 5)', fontsize=12)
    plt.ylim(0, 5)  # Ajustar el rango del eje Y a 0-5
    plt.title(titulo_grafico if titulo_grafico else 'Gráfico generado desde Python', fontsize=16)
    plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=len(encabezados), fontsize=10)  # Ajustar la leyenda

    # Guardar la gráfica como imagen
    import os
    output_path = os.path.join(os.path.dirname(ruta_excel), f"{titulo_grafico}.png")
    plt.savefig(output_path, bbox_inches='tight')
    plt.close()

# Función para cargar el archivo Excel usando tkinter
def cargar_archivo_excel():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de Tkinter
    ruta_archivo = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Archivos Excel", "*.xlsx")])
    return ruta_archivo

# Ejemplo de uso de la función
ruta_excel = cargar_archivo_excel()
if ruta_excel:
    generar_grafica_excel(ruta_excel, 'Hoja1')
