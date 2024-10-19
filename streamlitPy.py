#-*- coding: Latin1 -*-

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import io
import firebase_admin
from firebase_admin import credentials, db
import base64

# Configurar la p�gina para usar todo el ancho
st.set_page_config(layout='wide')

# Configurar la fuente globalmente
matplotlib.rcParams['font.family'] = 'Arial'

# Funci�n para inicializar Firebase
def inicializar_firebase():
    if not firebase_admin._apps:
        try:
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
                'databaseURL': 'https://excelgraph-67ab7-default-rtdb.europe-west1.firebasedatabase.app/'
            })
        except Exception as e:
            st.error(f"Error al inicializar Firebase: {e}")
            st.stop()

# Funci�n para sanitizar las claves (reemplaza caracteres no permitidos)
def sanitize_key(key):
    invalid_chars = ['.', '#', '$', '[', ']', '/']
    for char in invalid_chars:
        key = key.replace(char, '_')
    return key

# Funci�n para sanitizar las claves de forma recursiva en un diccionario o lista
def sanitize_keys(obj):
    if isinstance(obj, dict):
        new_obj = {}
        for key, value in obj.items():
            sanitized_key = sanitize_key(str(key))
            new_obj[sanitized_key] = sanitize_keys(value)
        return new_obj
    elif isinstance(obj, list):
        return [sanitize_keys(item) for item in obj]
    else:
        return obj

# Funci�n para cargar preferencias desde Firebase
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
        st.stop()

# Funci�n para guardar preferencias en Firebase
def guardar_preferencias(preferencias):
    try:
        ref = db.reference('preferencias')
        ref.set(preferencias)
    except Exception as e:
        st.error(f"Error al guardar preferencias: {e}")
        st.stop()

# Funci�n para generar enlace de descarga
def generar_enlace_descarga(data, filename, texto_link):
    b64 = base64.b64encode(data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">{texto_link}</a>'

# Definir colores por defecto para los valores (hardcodeados)
colores_valores_default = {
    '0': '#FF0000',  # Rojo
    '1': '#FFA500',  # Naranja
    '2': '#FFFF00',  # Amarillo
    '3': '#008000',  # Verde
    '4': '#0000FF',  # Azul
    '5': '#4B0082',  # �ndigo
}

# Definir colores por defecto para los rangos (hardcodeados)
colores_rangos_default = {
    '0-1': '#FFC0C0',  # Rojo claro
    '1-2': '#FFE0B2',  # Naranja claro
    '2-3': '#FFFFE0',  # Amarillo claro
    '3-4': '#C8E6C9',  # Verde claro
    '4-5': '#BBDEFB',  # Azul claro
}

# Inicializar Firebase
inicializar_firebase()

# Cargar las preferencias desde Firebase
preferencias_fijas = cargar_preferencias()

# T�tulo de la aplicaci�n
st.title("Generador de Gr�ficos Evolutivos desde Excel")

# Instrucciones para el usuario
st.write("""
    Esta aplicaci�n permite generar gr�ficos evolutivos a partir de un archivo de Excel.
    - **Descarga la plantilla de Excel y ed�tala con tus propias funciones y valores.**
    - Sube tu archivo Excel modificado.
    - Selecciona el tipo de gr�fico y proporciona un t�tulo.
    - Personaliza las series y los colores.
    - La aplicaci�n generar� el gr�fico con base en los datos proporcionados.
""")

# Proporcionar enlace para descargar la plantilla
st.write("### Descarga la plantilla de Excel:")
with open("plantilla.xlsx", "rb") as f:
    data = f.read()
    link = generar_enlace_descarga(data, "plantilla.xlsx", "Haz clic aqu� para descargar la plantilla")
    st.markdown(link, unsafe_allow_html=True)

# Cargar el archivo Excel
uploaded_file = st.file_uploader("Sube tu archivo de Excel", type=['xlsx'])

if uploaded_file is not None:
    # Leer el archivo Excel y extraer los datos
    df = pd.read_excel(uploaded_file, engine='openpyxl')

    # Obtener los encabezados de las columnas
    encabezados = df.columns.tolist()

    # Asegurarnos de que la primera columna es 'Funciones'
    if encabezados[0].lower() != 'funciones':
        st.error("El archivo Excel debe tener una columna llamada 'Funciones' como la primera columna.")
    else:
        # Obtener los nombres de las funciones
        funciones = df['Funciones'].astype(str).tolist()

        # Obtener las series (omitimos la columna 'Funciones')
        series = encabezados[1:]

        # Verificar que hay funciones y series
        if len(funciones) == 0 or len(series) == 0:
            st.error("El archivo Excel debe contener al menos una funci�n y una serie de datos.")
        else:
            # Crear un DataFrame con los datos
            df_datos = df.copy()

            # Inicializar el estado de sesi�n utilizando las preferencias fijas
            if 'colores_series' not in st.session_state:
                st.session_state['colores_series'] = preferencias_fijas.get('colores_series', {})
            if 'estilos_linea' not in st.session_state:
                st.session_state['estilos_linea'] = preferencias_fijas.get('estilos_linea', {})
            if 'grosor_linea' not in st.session_state:
                st.session_state['grosor_linea'] = preferencias_fijas.get('grosor_linea', {})
            if 'series_seleccionadas' not in st.session_state:
                st.session_state['series_seleccionadas'] = preferencias_fijas.get('series_seleccionadas', series)
            if 'colores_rangos' not in st.session_state:
                st.session_state['colores_rangos'] = preferencias_fijas.get('colores_rangos', {})
            if 'colores_valores' not in st.session_state:
                st.session_state['colores_valores'] = preferencias_fijas.get('colores_valores', {})

            # Barra lateral para los controles
            st.sidebar.header("Opciones de Personalizaci�n")

            # Solicitar el t�tulo del gr�fico
            titulo_grafico = st.sidebar.text_input("T�tulo del gr�fico", "Gr�fico generado desde Python")

            # Solicitar el tipo de gr�fico
            tipo_grafico = st.sidebar.selectbox("Tipo de gr�fico", ["L�nea", "Barra", "Dispersi�n"])

            # Seleccionar las series a incluir
            series_seleccionadas = st.sidebar.multiselect(
                "Selecciona las series a mostrar",
                options=series,
                default=st.session_state['series_seleccionadas']
            )
            st.session_state['series_seleccionadas'] = series_seleccionadas

            if not series_seleccionadas:
                st.error("Por favor, selecciona al menos una serie para mostrar en el gr�fico.")
            else:
                # Colores predeterminados
                colores_lineas_default = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

                # Personalizaci�n de colores y estilos
                colores_series = st.session_state['colores_series']
                estilos_linea = st.session_state['estilos_linea']
                grosor_linea = st.session_state['grosor_linea']

                for idx, serie in enumerate(series_seleccionadas):
                    with st.sidebar.expander(f"Personalizar serie: {serie}", expanded=False):
                        # Color predeterminado
                        color_default = colores_series.get(serie, colores_lineas_default[idx % len(colores_lineas_default)])
                        color = st.color_picker(f"Color para {serie}", color_default)
                        colores_series[serie] = color

                        # Estilo de l�nea predeterminado
                        estilo_default = estilos_linea.get(serie, 'S�lida')
                        estilo = st.selectbox(
                            f"Estilo de l�nea para {serie}",
                            ['S�lida', 'Punteada', 'Punteada-Punteada'],
                            index=['S�lida', 'Punteada', 'Punteada-Punteada'].index(estilo_default)
                        )
                        estilos_linea[serie] = estilo

                        # Grosor de l�nea predeterminado
                        grosor_default = grosor_linea.get(serie, 3)
                        grosor = st.slider(f"Grosor de l�nea para {serie}", 1, 5, grosor_default)
                        grosor_linea[serie] = grosor

                # Guardar en el estado de sesi�n
                st.session_state['colores_series'] = colores_series
                st.session_state['estilos_linea'] = estilos_linea
                st.session_state['grosor_linea'] = grosor_linea

                # Personalizaci�n de los colores de los valores
                colores_valores = st.session_state['colores_valores']
                with st.sidebar.expander("Personalizar colores de valores", expanded=False):
                    for valor in sorted(colores_valores_default.keys()):
                        # Aseguramos que valor es una cadena
                        valor_str = str(valor)
                        color_default = colores_valores.get(valor_str, colores_valores_default[valor_str])
                        color = st.color_picker(f"Color para valor {valor}", color_default)
                        colores_valores[valor_str] = color  # Guardar como cadena para serializaci�n JSON

                st.session_state['colores_valores'] = colores_valores

                # Personalizaci�n de los colores de fondo de los rangos
                colores_rangos = st.session_state['colores_rangos']
                with st.sidebar.expander("Personalizar colores de fondo de rangos", expanded=False):
                    for rango in colores_rangos_default.keys():
                        color_default = colores_rangos.get(rango, colores_rangos_default[rango])
                        color = st.color_picker(f"Color para rango {rango}", color_default)
                        colores_rangos[rango] = color

                st.session_state['colores_rangos'] = colores_rangos

                # Mapear estilos de l�nea a formatos de Matplotlib
                estilos_mpl = {
                    'S�lida': '-',
                    'Punteada': '--',
                    'Punteada-Punteada': ':'
                }

                # Crear la gr�fica con Matplotlib
                fig, ax = plt.subplots(figsize=(25, 9), dpi=300)

                # Seleccionar el tipo de gr�fico
                if tipo_grafico.lower() == 'barra':
                    width = 0.8 / len(series_seleccionadas)  # Ancho de cada barra
                    x = range(len(funciones))
                    for idx, serie in enumerate(series_seleccionadas):
                        ax.bar(
                            [pos + idx * width for pos in x],
                            df_datos[serie],
                            width=width,
                            color=colores_series.get(serie, colores_lineas_default[idx % len(colores_lineas_default)]),
                            label=serie
                        )
                    ax.set_xticks([pos + width * (len(series_seleccionadas) - 1) / 2 for pos in x])
                    ax.set_xticklabels(funciones)
                elif tipo_grafico.lower() == 'dispersi�n':
                    for idx, serie in enumerate(series_seleccionadas):
                        ax.scatter(
                            funciones,
                            df_datos[serie],
                            color=colores_series.get(serie, colores_lineas_default[idx % len(colores_lineas_default)]),
                            label=serie
                        )
                else:
                    for idx, serie in enumerate(series_seleccionadas):
                        ax.plot(
                            funciones,
                            df_datos[serie],
                            marker='o',
                            color=colores_series.get(serie, colores_lineas_default[idx % len(colores_lineas_default)]),
                            linestyle=estilos_mpl.get(estilos_linea.get(serie, 'S�lida'), '-'),
                            linewidth=grosor_linea.get(serie, 3),
                            markerfacecolor='white',
                            markeredgecolor=colores_series.get(serie, colores_lineas_default[idx % len(colores_lineas_default)]),
                            label=serie
                        )

                # Rellenar el fondo del gr�fico en base a los colores de los rangos
                for rango in colores_rangos_default.keys():
                    y0, y1 = map(float, rango.split('-'))
                    color = colores_rangos.get(rango, colores_rangos_default[rango])
                    ax.axhspan(y0, y1, facecolor=color, alpha=0.3)

                # Cambiar el color del texto de los r�tulos del eje X en base a los valores de la primera serie seleccionada
                primera_serie = series_seleccionadas[0]
                valores_serie = df_datos[primera_serie].tolist()
                colores_rotulos = []
                for valor in valores_serie:
                    if pd.notna(valor) and isinstance(valor, (int, float)):
                        valor_str = str(int(valor))
                        color_rotulo = colores_valores.get(valor_str, '#000000')
                    else:
                        color_rotulo = '#000000'
                    colores_rotulos.append(color_rotulo)

                # Ajustar los r�tulos del eje X
                etiquetas = [
                    funcion if isinstance(funcion, str) else ''
                    for funcion in funciones
                ]

                plt.xticks(
                    ticks=range(len(funciones)),
                    labels=etiquetas,
                    fontsize=8,
                    fontweight='bold',
                    ha='right',
                    rotation=15  # Rotar las etiquetas 15 grados
                )

                for tick_label, color in zip(ax.get_xticklabels(), colores_rotulos):
                    tick_label.set_color(color)
                    tick_label.set_fontweight('bold')
                    tick_label.set_antialiased(True)

                plt.xlabel('Funciones', fontsize=12, fontweight='bold')
                plt.ylabel('Valores', fontsize=12, fontweight='bold')
                plt.title(titulo_grafico, fontsize=16, fontweight='bold')
                plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=len(series_seleccionadas), fontsize=10)

                # Ajustar el rango del eje Y
                plt.ylim(0, 5)

                plt.tight_layout()

                # Mostrar la gr�fica con Streamlit usando un buffer para mayor nitidez
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=300)
                buf.seek(0)
                st.image(buf, use_column_width=True)

                # Opci�n para descargar la gr�fica generada en formato PNG y SVG
                st.write("### Descargar gr�fico:")
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

                # Bot�n para guardar preferencias fijas en Firebase
                if st.sidebar.button("Guardar preferencias fijas"):
                    preferencias_actualizadas = {
                        'colores_series': st.session_state['colores_series'],
                        'estilos_linea': st.session_state['estilos_linea'],
                        'grosor_linea': st.session_state['grosor_linea'],
                        'series_seleccionadas': st.session_state['series_seleccionadas'],
                        'colores_rangos': st.session_state['colores_rangos'],
                        'colores_valores': st.session_state['colores_valores']
                    }

                    # Sanitizar las claves antes de guardar
                    preferencias_sanitizadas = sanitize_keys(preferencias_actualizadas)

                    # Guardar las preferencias sanitizadas
                    guardar_preferencias(preferencias_sanitizadas)
                    st.sidebar.success("Preferencias guardadas correctamente.")
