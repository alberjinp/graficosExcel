#-*- coding: Latin1 -*-


import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import matplotlib
import io
import json
import firebase_admin
from firebase_admin import credentials, db

# Inicializar Firebase
def inicializar_firebase():
    if not firebase_admin._apps:
        cred = credentials.Certificate({
            "type": st.secrets["firebase"]["type"],
            "project_id": st.secrets["firebase"]["project_id"],
            "private_key_id": st.secrets["firebase"]["private_key_id"],
            "private_key": st.secrets["firebase"]["private_key"],
            "client_email": st.secrets["firebase"]["client_email"],
            "client_id": st.secrets["firebase"]["client_id"],
            "auth_uri": st.secrets["firebase"]["auth_uri"],
            "token_uri": st.secrets["firebase"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["firebase"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url": st.secrets["firebase"]["client_x509_cert_url"]
        })
        firebase_admin.initialize_app(cred, {
            'databaseURL': 'https://tu_project_id-default-rtdb.firebaseio.com/'
        })

# Función para cargar preferencias desde Firebase
def cargar_preferencias():
    try:
        ref = db.reference('preferencias')
        preferencias = ref.get()
        if preferencias:
            return preferencias
        else:
            return {}
    except Exception as e:
        st.error(f"Error al cargar preferencias: {e}")
        return {}

# Función para guardar preferencias en Firebase
def guardar_preferencias(preferencias):
    try:
        ref = db.reference('preferencias')
        ref.set(preferencias)
    except Exception as e:
        st.error(f"Error al guardar preferencias: {e}")

# Inicializar Firebase
inicializar_firebase()

# Cargar las preferencias desde Firebase
preferencias_fijas = cargar_preferencias()

# Configurar la página para usar todo el ancho
st.set_page_config(layout='wide')

# Configurar la fuente globalmente
matplotlib.rcParams['font.family'] = 'Arial'

# Título de la aplicación
st.title("Generador de Gráficos Evolutivos desde Excel")

# Instrucciones para el usuario
st.write("""
    Esta aplicación permite generar gráficos evolutivos a partir de un archivo de Excel.
    - Sube tu archivo Excel.
    - Selecciona el tipo de gráfico y proporciona un título.
    - Personaliza las series y los colores.
    - La aplicación generará el gráfico con base en los datos proporcionados.
""")

# Cargar el archivo Excel
uploaded_file = st.file_uploader("Sube tu archivo de Excel", type=['xlsx'])

if uploaded_file is not None:
    # Leer el archivo Excel y la hoja específica
    wb = openpyxl.load_workbook(uploaded_file)
    hoja_nombre = 'Hoja1'
    hoja = wb[hoja_nombre]

    df = pd.read_excel(uploaded_file, sheet_name=hoja_nombre, header=None, usecols="A:F", engine='openpyxl')

    # Obtener los encabezados de la fila 5 (Basal, Clínica, Funcionalidad, etc.)
    encabezados = df.iloc[4, 1:].tolist()

    # Obtener los nombres de las funciones desde A6 hasta A31
    funciones = df.iloc[5:31, 0].astype(str).tolist()

    # Obtener los valores de las series desde B6:F31
    datos = df.iloc[5:31, 1:].values

    # Crear un DataFrame para facilitar la manipulación
    df_datos = pd.DataFrame(datos, columns=encabezados)
    df_datos.insert(0, 'Funciones', funciones)

    # Inicializar el estado de sesión utilizando las preferencias fijas
    if 'colores_series' not in st.session_state:
        st.session_state['colores_series'] = preferencias_fijas.get('colores_series', {})
    if 'estilos_linea' not in st.session_state:
        st.session_state['estilos_linea'] = preferencias_fijas.get('estilos_linea', {})
    if 'grosor_linea' not in st.session_state:
        st.session_state['grosor_linea'] = preferencias_fijas.get('grosor_linea', {})
    if 'series_seleccionadas' not in st.session_state:
        st.session_state['series_seleccionadas'] = preferencias_fijas.get('series_seleccionadas', encabezados)
    if 'colores_rangos' not in st.session_state:
        st.session_state['colores_rangos'] = preferencias_fijas.get('colores_rangos', {})
    if 'colores_valores' not in st.session_state:
        st.session_state['colores_valores'] = preferencias_fijas.get('colores_valores', {})

    # Barra lateral para los controles
    st.sidebar.header("Opciones de Personalización")

    # Solicitar el título del gráfico
    titulo_grafico = st.sidebar.text_input("Título del gráfico", "Gráfico generado desde Python")

    # Solicitar el tipo de gráfico
    tipo_grafico = st.sidebar.selectbox("Tipo de gráfico", ["Línea", "Barra", "Dispersión"])

    # Seleccionar las series a incluir
    series_seleccionadas = st.sidebar.multiselect(
        "Selecciona las series a mostrar",
        options=encabezados,
        default=st.session_state['series_seleccionadas']
    )
    st.session_state['series_seleccionadas'] = series_seleccionadas

    if not series_seleccionadas:
        st.error("Por favor, selecciona al menos una serie para mostrar en el gráfico.")
    else:
        # Colores predeterminados
        colores_lineas_default = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

        # Personalización de colores y estilos
        colores_series = st.session_state['colores_series']
        estilos_linea = st.session_state['estilos_linea']
        grosor_linea = st.session_state['grosor_linea']

        for idx, serie in enumerate(series_seleccionadas):
            with st.sidebar.expander(f"Personalizar serie: {serie}", expanded=False):
                # Color predeterminado
                color_default = colores_series.get(serie, colores_lineas_default[idx % len(colores_lineas_default)])
                color = st.color_picker(f"Color para {serie}", color_default)
                colores_series[serie] = color

                # Estilo de línea predeterminado
                estilo_default = estilos_linea.get(serie, 'Sólida')
                estilo = st.selectbox(
                    f"Estilo de línea para {serie}",
                    ['Sólida', 'Punteada', 'Punteada-Punteada'],
                    index=['Sólida', 'Punteada', 'Punteada-Punteada'].index(estilo_default)
                )
                estilos_linea[serie] = estilo

                # Grosor de línea predeterminado
                grosor_default = grosor_linea.get(serie, 3)
                grosor = st.slider(f"Grosor de línea para {serie}", 1, 5, grosor_default)
                grosor_linea[serie] = grosor

        # Guardar en el estado de sesión
        st.session_state['colores_series'] = colores_series
        st.session_state['estilos_linea'] = estilos_linea
        st.session_state['grosor_linea'] = grosor_linea

        # Obtener los valores y colores desde B34:D39 para los colores de los valores
        colores_valores_default = {}
        for fila in range(34, 40):
            valor = hoja[f'B{fila}'].value
            color_hex = hoja[f'D{fila}'].value
            if not color_hex:
                color_hex = '#000000'
            colores_valores_default[valor] = color_hex

        # Personalización de los colores de los valores
        colores_valores = st.session_state['colores_valores']
        with st.sidebar.expander("Personalizar colores de valores", expanded=False):
            for valor in sorted(colores_valores_default.keys()):
                color_default = colores_valores.get(str(valor), colores_valores_default[valor])
                color = st.color_picker(f"Color para valor {valor}", color_default)
                colores_valores[str(valor)] = color  # Guardar como cadena para serialización JSON

        st.session_state['colores_valores'] = colores_valores

        # Personalización de los colores de fondo de los rangos
        colores_rangos_default = {
            '0-1': colores_valores_default.get(0, '#FFFFFF'),
            '1-2': colores_valores_default.get(1, '#FFFFFF'),
            '2-3': colores_valores_default.get(2, '#FFFFFF'),
            '3-4': colores_valores_default.get(3, '#FFFFFF'),
            '4-5': colores_valores_default.get(4, '#FFFFFF'),
        }

        colores_rangos = st.session_state['colores_rangos']
        with st.sidebar.expander("Personalizar colores de fondo de rangos", expanded=False):
            for rango in ['0-1', '1-2', '2-3', '3-4', '4-5']:
                color_default = colores_rangos.get(rango, colores_rangos_default[rango])
                color = st.color_picker(f"Color para rango {rango}", color_default)
                colores_rangos[rango] = color

        st.session_state['colores_rangos'] = colores_rangos

        # Mapear estilos de línea a formatos de Matplotlib
        estilos_mpl = {
            'Sólida': '-',
            'Punteada': '--',
            'Punteada-Punteada': ':'
        }

        # Crear la gráfica con Matplotlib
        fig, ax = plt.subplots(figsize=(22, 9), dpi=300)

        # Seleccionar el tipo de gráfico
        if tipo_grafico.lower() == 'barra':
            width = 0.8 / len(series_seleccionadas)  # Ancho de cada barra
            x = range(len(funciones))
            for idx, serie in enumerate(series_seleccionadas):
                ax.bar(
                    [pos + idx * width for pos in x],
                    df_datos[serie],
                    width=width,
                    color=colores_series[serie],
                    label=serie
                )
            ax.set_xticks([pos + width * (len(series_seleccionadas) - 1) / 2 for pos in x])
            ax.set_xticklabels(funciones)
        elif tipo_grafico.lower() == 'dispersión':
            for serie in series_seleccionadas:
                ax.scatter(
                    funciones,
                    df_datos[serie],
                    color=colores_series[serie],
                    label=serie
                )
        else:
            for serie in series_seleccionadas:
                ax.plot(
                    funciones,
                    df_datos[serie],
                    marker='o',
                    color=colores_series[serie],
                    linestyle=estilos_mpl[estilos_linea[serie]],
                    linewidth=grosor_linea[serie],
                    markerfacecolor='white',
                    markeredgecolor=colores_series[serie],
                    label=serie
                )

        # Rellenar el fondo del gráfico en base a los colores de los rangos
        for rango in ['0-1', '1-2', '2-3', '3-4', '4-5']:
            y0, y1 = map(float, rango.split('-'))
            color = colores_rangos[rango]
            ax.axhspan(y0, y1, facecolor=color, alpha=0.3)

        # Cambiar el color del texto de los rótulos del eje X en base a los valores de Basal
        valores_basal = df.iloc[5:31, 1].tolist()
        colores_rotulos = [
            colores_valores.get(str(valor), '#000000') if pd.notna(valor) and isinstance(valor, (int, float))
            else '#000000' for valor in valores_basal
        ]

        # Ajustar los rótulos del eje X
        etiquetas = [
            funcion.replace(' ', '\n') if isinstance(funcion, str) else ''
            for funcion in funciones
        ]

        plt.xticks(
            ticks=range(len(funciones)),
            labels=etiquetas,
            fontsize=8,
            fontweight='bold',
            ha='center'
        )

        for tick_label, color in zip(ax.get_xticklabels(), colores_rotulos):
            tick_label.set_color(color)
            tick_label.set_fontweight('bold')
            tick_label.set_antialiased(True)

        plt.xlabel('Funciones', fontsize=12, fontweight='bold')
        plt.ylabel('Valores (0 a 5)', fontsize=12, fontweight='bold')
        plt.ylim(0, 5)  # Ajustar el rango del eje Y a 0-5
        plt.title(titulo_grafico, fontsize=16, fontweight='bold')
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=len(series_seleccionadas), fontsize=10)

        plt.tight_layout()

        # Mostrar la gráfica con Streamlit usando un buffer para mayor nitidez
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=300)
        buf.seek(0)
        st.image(buf, use_column_width=True)

        # Opción para descargar la gráfica generada en formato PNG y SVG
        st.write("### Descargar gráfico:")
        col1, col2 = st.columns(2)

        with col1:
            output_path_png = f"{titulo_grafico}.png"
            fig.savefig(output_path_png, format='png', dpi=300, bbox_inches='tight')
            with open(output_path_png, "rb") as file:
                btn_png = st.download_button(
                    label="Descargar PNG",
                    data=file,
                    file_name=output_path_png,
                    mime="image/png"
                )
            st.write("**PNG**: Adecuado para uso en pantallas y documentos digitales.")

        with col2:
            output_path_svg = f"{titulo_grafico}.svg"
            fig.savefig(output_path_svg, format='svg', bbox_inches='tight')
            with open(output_path_svg, "rb") as file:
                btn_svg = st.download_button(
                    label="Descargar SVG",
                    data=file,
                    file_name=output_path_svg,
                    mime="image/svg+xml"
                )
            st.write("**SVG**: Ideal para escalado y alta calidad en impresiones.")

        # Botón para guardar preferencias fijas en Firebase
        if st.sidebar.button("Guardar preferencias fijas"):
            preferencias_actualizadas = {
                'colores_series': st.session_state['colores_series'],
                'estilos_linea': st.session_state['estilos_linea'],
                'grosor_linea': st.session_state['grosor_linea'],
                'series_seleccionadas': st.session_state['series_seleccionadas'],
                'colores_rangos': st.session_state['colores_rangos'],
                'colores_valores': st.session_state['colores_valores']
            }
            guardar_preferencias(preferencias_actualizadas)
            st.sidebar.success("Preferencias guardadas correctamente.")
