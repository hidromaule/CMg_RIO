import datetime
import os
import re
from pathlib import Path
import cloudscraper
import zipfile
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time

st.set_page_config(layout="wide")
col1, col2, col3 = st.columns([1, 6, 1])

#### Inicialización de variable de sesión para actualizar datos ####
if "actualizar" not in st.session_state:
    st.session_state.actualizar = True

#### Botón para actualizar datos ####
with col3:
    if st.button("🔄 Actualizar datos"):
        st.session_state.actualizar = True

st.markdown(
    """
    <style>
    img {
        border-radius: 0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

with col1:
    st.image("Assets/logo.png",width=180)

#### Función para esperar a que la descarga se complete ####
def wait_for_download(folder, timeout=60):
    flag = "Ok"
    for _ in range(timeout):
        files = os.listdir(folder)
        if files and not any(f.endswith(".crdownload") for f in files):
            return flag
        time.sleep(1)
    raise TimeoutError("Descarga no completada")

#### Función para descargar archivos protegidos por Cloudflare ####
def download_zip_cloudflare(url: str, output_path: str, archivo: str):
    scraper = cloudscraper.create_scraper(
        browser={
            "browser": "chrome",
            "platform": "windows",
            "desktop": True
        }
    )

    r = scraper.get(url, stream=True)
    r.raise_for_status()

    with open(output_path, "wb") as f:
        for chunk in r.iter_content(8192):
            f.write(chunk)

    #print("Se descargó el archivo:", archivo)

#### Función para descargar archivos usando selenium ####
def selenium_download(download_path,file_name):
    #### Se agregan opciones de Chrome, para funciona en modo headless ####
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )

    #### Se configuran y agregan las preferencias de descarga ####
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    #### Se inicializa el driver de Chrome ####
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    #### Se abre la página de descarga ####
    driver.get("https://programa.coordinador.cl/operacion/pcp/bases-modelo")

    #### Se espera a que la tabla de datos cargue ####
    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//td[starts-with(normalize-space(), 'PROGRAMA')]")
        )
    )

    #### Se busca el botón ZIP asociado al archivo correcto ####
    zip_button = wait.until(
        EC.presence_of_element_located((
            By.XPATH,
            "//td[normalize-space()='" + file_name + "']/parent::tr//a[contains(@class,'descargarBtn')]"
        ))
    )

    #### Baja y se forzarfuerza el foco en la fila buscada ####
    driver.execute_script("""
        arguments[0].scrollIntoView({block:'center'});
        arguments[0].focus();
    """, zip_button)

    #### Pequeño delay para estabilidad de Angular ####
    time.sleep(0.5)

    #### Click en el botón de descarga ####
    driver.execute_script("arguments[0].click();", zip_button)

    #### Se espera 3 segundos para cerrar el navegador ####
    time.sleep(3)

    #### Se espera a que termine la descarga y se levanta un flag ####
    flag = wait_for_download(download_path)
    return flag

#### Función para extraer archivos ZIP ####
def extract_single_file(zip_filename, file_to_extract, extract_path='.'):
    # Revisar si la ruta de extracción existe, si no, crearla
    if not os.path.exists(extract_path):
        os.makedirs(extract_path)

    # Abrir el archivo ZIP y extraer el archivo específico
    try:
        with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
            # Extraer el archivo específico
            zip_ref.extract(file_to_extract, extract_path)
            pass #print(f"Se extrajo el archivo '{file_to_extract}' en '{extract_path}'")
    except KeyError:
        pass #print(f"Error: '{file_to_extract}' no encontrado en el archivo ZIP.")
    except zipfile.BadZipFile:
        pass #print(f"Error: '{zip_filename}' no es un archivo ZIP válido.")
    except RuntimeError as e:
        pass #print(f"Error extrayendo el archivo: {e}") #Maneja errores como archvios encriptados

#### Función para obtener el bloque horario correspondiente ####
def obtener_bloque(hora):
    h = int(hora.split(":")[0])

    if 0 <= h < 8:
        return bloque_1
    elif 8 <= h < 16:
        return bloque_2
    else:
        return bloque_3

#### Función para limpiar archivos antiguos en la carpeta ####
def limpiar_archivos_antiguos(carpeta, dry_run=True):
    hoy = datetime.datetime.today()

    patron_energia = re.compile(r"^ENERGIA(\d{8})\.csv$")
    patron_poa = re.compile(r"^PO(\d{6})\.xlsx$")

    for archivo in os.listdir(carpeta):
        ruta = os.path.join(carpeta, archivo)

        if not os.path.isfile(ruta):
            continue

        fecha = None

        m1 = patron_energia.match(archivo)
        m2 = patron_poa.match(archivo)

        if m1:
            fecha = datetime.datetime.strptime(m1.group(1), "%Y%m%d")
        elif m2:
            fecha = datetime.datetime.strptime(m2.group(1), "%y%m%d")

        if fecha is None:
            continue

        if fecha.date() < hoy.date():
            if dry_run:
                pass #print(f"[DRY RUN] Se eliminaría el archivo: {archivo}")
            else:
                os.remove(ruta)
                pass #print(f"Se eliminó el archivo: {archivo}")

#### Función para convertir hora (HH:MM o HH:MM:SS) a datetime ####
def hora_a_datetime(hora_str):
    partes = hora_str.split(":")

    if len(partes) == 2:        # HH:MM
        fmt = "%H:%M"
    elif len(partes) == 3:      # HH:MM:SS
        fmt = "%H:%M:%S"
    else:
        raise ValueError(f"Formato de hora no reconocido: {hora_str}")

    return datetime.datetime.combine(
        datetime.date.today(),
        datetime.datetime.strptime(hora_str, fmt).time()
    )

#### Función para agregar sombreado de prorratas en gráfico de Plotly ####
def agregar_sombreado_prorratas_plotly(fig, lista_CMg, resolucion_min=15):
    comentario_actual = None
    x_inicio = None

    for fila in lista_CMg:
        x = hora_a_datetime(fila[0])
        comentario = fila[2]

        if comentario != comentario_actual:
            # cerrar tramo anterior
            if comentario_actual in (
                "Prorrata Generalizada",
                "Prorrata Generalizada costo SEN 0"
            ):
                color = (
                    "rgba(128, 0, 128, 0.25)"   # morado
                    if comentario_actual == "Prorrata Generalizada"
                    else "rgba(255, 255, 0, 0.35)"  # amarillo
                )

                fig.add_vrect(
                    x0=x_inicio,
                    x1=x,
                    fillcolor=color,
                    line_width=0,
                    layer="below"
                )

            comentario_actual = comentario
            x_inicio = x

    # cerrar último tramo
    if comentario_actual in (
        "Prorrata Generalizada",
        "Prorrata Generalizada costo SEN 0"
    ):
        x_fin = hora_a_datetime(lista_CMg[-1][0])
        x_fin = x_fin + datetime.timedelta(minutes=resolucion_min)

        color = (
            "rgba(128, 0, 128, 0.25)"
            if comentario_actual == "Prorrata Generalizada"
            else "rgba(255, 255, 0, 0.35)"
        )

        fig.add_vrect(
            x0=x_inicio,
            x1=x_fin,
            fillcolor=color,
            line_width=0,
            layer="below",
            #hoverinfo="skip"
        )

#### Se obtiene la fecha de hoy ####
today = datetime.date.today()
year = today.year

#### Se generan las URLs de descarga del programa y RIOs ####
URL1 = "https://www.coordinador.cl/wp-content/uploads/" + str(year) + "/" + today.strftime("%m") + "/PROGRAMA" + str(year) + today.strftime("%m") + today.strftime("%d") + ".zip"
URL2 = "https://www.coordinador.cl/wp-admin/admin-ajax.php?action=export_energia_csv&fecha_inicio=" + str(today) + "&fecha_termino=" + str(today) + "&hora_inicio=00:00:00&hora_termino=23:59:59"

#### Se crea una carpeta temporal para los archivos ####
carpeta = "BD"
BASE_DIR = Path(__file__).resolve().parent
output_dir = BASE_DIR / carpeta
output_dir.mkdir(exist_ok=True)

if st.session_state.actualizar:
    
    with col2:
        with st.spinner("Descargando y procesando datos..."):
            print(str(datetime.datetime.now()) + " - Actualizando datos...")
            
            #### Se limpian los archivos antiguos en la carpeta ####
            limpiar_archivos_antiguos(output_dir, dry_run=False)

            #### Se generan los nombres de los archivos ####
            archivo1 = "PROGRAMA" + str(year) + today.strftime("%m") + today.strftime("%d") + ".zip"
            archivo2 = "ENERGIA" + str(year) + today.strftime("%m") + today.strftime("%d") + ".csv"
            file_path1 = output_dir / archivo1
            file_path2 = output_dir / archivo2


            #### Se chequea si los archivos ya existen antes de descargarlos y extraerlos ####
            if os.path.isfile(file_path1) or os.path.isfile(output_dir / ("PO" + today.strftime("%y") + today.strftime("%m") + today.strftime("%d") + ".xlsx")):
                pass #print("El archivo PO ya existe. No se descargará de nuevo.")
            else:
                flag = selenium_download(str(output_dir),archivo1)
                if flag != "Ok":
                    print("Error en la descarga del archivo PO.")    

            download_zip_cloudflare(URL2, file_path2, archivo2)

            if os.path.isfile(output_dir / ("PO" + today.strftime("%y") + today.strftime("%m") + today.strftime("%d") + ".xlsx")):
                pass #print("El archivo PO" + today.strftime("%y") + today.strftime("%m") + today.strftime("%d") + ".xlsx ya existe. No se extraerá de nuevo.")
            else:
                extract_single_file(file_path1, "PO" + today.strftime("%y") + today.strftime("%m") + today.strftime("%d") + ".xlsx", output_dir)

            #### Se elimina el archivo ZIP descargado para ahorrar espacio ####
            if os.path.exists(file_path1):
                os.remove(file_path1)
                pass #print(f"Se eliminó el archivo {archivo1}")
            else:
                pass #print(f"El archivo {archivo1} no existe.")

            #### Se generan las rutas a los archivos de datos ####
            PO = output_dir / ("PO" + today.strftime("%y") + today.strftime("%m") + today.strftime("%d") + ".xlsx")
            energia = file_path2

            #### Se genera el DataFrame de energía, se eliminan las 3 primeras filas y las columnas innecesarias ####
            df_energia = pd.read_csv(
                energia,
                sep=";",
                skiprows=4,
                header=0
            )
            df_energia = df_energia.drop(columns=["FECHA","NOMBRE CONFIGURACIÓN","UNIDAD GENERADORA","POTENCIA MÁXIMA","POTENCIA MÍNIMA","POTENCIA INSTRUIDA","ESTADO OPERACIONAL","ESTADO OPERACIONAL COMBUSTIBLE","CONSIGNAS","CONSIGNA LIMITACIÓN","MOTIVO","SENTIDO FLUJO","ESTADO DE EMBALSE","Nº DOCUMENTO","CENTRO DE CONTROL","Fecha de Edición Registro"],axis=1)

            #### Se transforma el DataFrame de energía en una lista de listas ####
            lista_energia = df_energia.values.tolist()

            #### Se genera el DataFrame del PO se eliminan las 6 primeras filas y las columnas innecesarias ####
            df_PO = pd.read_excel(PO, sheet_name="TCO", skiprows=6)
            df_PO = df_PO.drop(df_PO.columns[[0,1,4,5,8,9]],axis=1)

            #### Se generan los diccionarios de CMg por bloque ####
            bloque_1 = dict(zip(df_PO["CENTRALES"].dropna(), df_PO["CMg [USD/MWh]"].dropna()))
            bloque_2 = dict(zip(df_PO["CENTRALES.1"].dropna(), df_PO["CMg [USD/MWh].1"].dropna()))
            bloque_3 = dict(zip(df_PO["CENTRALES.2"].dropna(), df_PO["CMg [USD/MWh].2"].dropna()))

            #### Se genera una lista auxiliar y se cruza con los diccionarios de costo marginal ####
            lista_CMg = []

            for fila in lista_energia:
                nueva_fila = fila.copy()
                hora = fila[0]

                bloque = obtener_bloque(hora)

                for i in range(3, 12):  # columnas 4 a 12
                    nombre = fila[i]

                    if nombre == "ERNC":
                        nueva_fila[i] = 0
                    else:
                        nueva_fila[i] = bloque.get(nombre, None)  
                        # .get evita error si el nombre no existe

                lista_CMg.append(nueva_fila)

            #### Se invierte la lista para que vaya de 00:00 a 23:59 ####
            lista_CMg = lista_CMg[::-1]

            #### Se rellenan los valores None con el valor anterior ####
            for i in range(len(lista_CMg)):
                for k in range(3, 12):
                    if lista_CMg[i][k] is None:
                        lista_CMg[i][k] = lista_CMg[i-1][k]
        
    st.session_state.actualizar = False
    print(str(datetime.datetime.now()) + " - Datos actualizados.")

#### Se ocultan los botones de la esquina superior derecha ####
st.markdown(
    """
    <style>
        [data-testid="stToolbar"] {display: none;}
    </style>
    """,
    unsafe_allow_html=True
)

with col3:
    st.caption(f"Última actualización: {datetime.datetime.now():%H:%M:%S}")

#### Nombres de las barras ####
nombres = [
    "Crucero 220", "Diego de Almagro 220", "Cardones 220",
    "Pan de Azucar 220", "Las Palmas 220", "Quillota 220",
    "Alto Jahuel 220", "Charrúa 220", "Puerto Montt 220"
]

#### Parámetros visuales del gráfico ####
font_title = 24
font_axes = 17
font_ticks = 13
font_legend = 13
height = 600

#### Eje x para el gráfico en plotly####
x_datetime = [hora_a_datetime(fila[0]) for fila in lista_CMg]

fig = go.Figure()

#### Se grafican las líneas de costo marginal ####
for i, nombre in zip(range(3, 12), nombres):
    valores = [fila[i] for fila in lista_CMg]

    fig.add_trace(
        go.Scatter(
            x=x_datetime,
            y=valores,
            mode="lines+markers",
            name=nombre,
            hovertemplate=
                "<b>%{fullData.name}</b><br>" +
                "Costo: %{y:,.1f}<extra></extra>"
        )
    )

#### Se sombrean los horarios de prorratas ####
agregar_sombreado_prorratas_plotly(fig, lista_CMg)

#### Se crean lineas fantasmas para la leyenda del sombreado ####
fig.add_trace(
    go.Scatter(
        x=[None],
        y=[None],
        mode="markers",
        marker=dict(
            size=12,
            color="rgba(128, 0, 128, 0.25)",
            symbol="square"
        ),
        name="Prorrata Generalizada (rampas)",
        showlegend=True,
        hoverinfo="skip"
    )
)

fig.add_trace(
    go.Scatter(
        x=[None],
        y=[None],
        mode="markers",
        marker=dict(
            size=12,
            color="rgba(255, 255, 0, 0.25)",
            symbol="square"
        ),
        name="Prorrata Generalizada Costo 0",
        showlegend=True,
        hoverinfo="skip"
    )
)

#### Configuración de ejes y layout ####
fig.update_xaxes(
    title="Hora",
    tickformat="%H:%M",
    dtick=60 * 60 * 1000,
    title_font=dict(size=font_axes),
    tickfont=dict(size=font_ticks)
)

fig.update_yaxes(
    dtick=25,
    title_font=dict(size=font_axes),
    tickfont=dict(size=font_ticks)
)

fig.update_layout(
    title=dict(
        text="Costos marginales SEN  -  Registro de Instrucciones de Operación  -  " + today.strftime("%d-%m-%Y"),
        x=0.5,
        xanchor="center",
        font=dict(
            size=font_title,
            family="Arial",
            color="black"
        )
    ),
    yaxis_title="CMg (USD/MWh)",
    hovermode="x unified",
    xaxis=dict(
        unifiedhovertitle=dict(text='<b>%{x|%H:%M}</b>')
    ),
    legend=dict(
        orientation="h",
        x=0.5,
        xanchor="center",
        y=-0.30,
        yanchor="top",
        itemwidth=30,
        font=dict(
            size=font_legend
        ),
    ),
    separators=",.",
    height=height
)

#### Se muestra el gráfico ####
with col2:
    st.plotly_chart(fig)