# CREACIÓN DE DASHBOARD PARA LA PRESENTACIÓN DE RESULTADOS

import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd
import plotly.express as px
import subprocess
import os
import psutil  # Para verificar si el servidor está ejecutándose
import time
from datetime import datetime
import platform

# Función para encontrar 'soffice' en ubicaciones comunes
def find_soffice_executable():
    import platform

    system = platform.system()

    possible_paths = []

    if system == 'Darwin':  # macOS
        possible_paths = [
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            '/Applications/OpenOffice.app/Contents/MacOS/soffice',
        ]
    elif system == 'Linux':
        possible_paths = [
            '/usr/bin/soffice',
            '/usr/local/bin/soffice',
            '/snap/bin/soffice',
        ]
    elif system == 'Windows':
        possible_paths = [
            os.path.join(os.environ.get('PROGRAMFILES', 'C:\\Program Files'), 'LibreOffice', 'program', 'soffice.exe'),
            os.path.join(os.environ.get('PROGRAMFILES(X86)', 'C:\\Program Files (x86)'), 'LibreOffice', 'program', 'soffice.exe'),
        ]
    else:
        raise EnvironmentError(f"Sistema operativo no soportado: {system}")

    for path in possible_paths:
        if os.path.isfile(path):
            return path

    # Si no se encontró, intentar obtener desde variable de entorno
    soffice_path = os.environ.get('SOFFICE_PATH')
    if soffice_path and os.path.isfile(soffice_path):
        return soffice_path

    # Si aún no se encontró, lanzar un error
    raise FileNotFoundError("No se pudo encontrar 'soffice'. Por favor, instala LibreOffice o especifica la ruta al ejecutable en la variable de entorno 'SOFFICE_PATH'.")

# Función para verificar si LibreOffice está ejecutándose
def is_libreoffice_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'soffice.bin':
            return True
    return False

# Función para iniciar el servidor de LibreOffice
def start_libreoffice_server():
    soffice_executable = find_soffice_executable()
    subprocess.Popen([soffice_executable, '--headless', '--accept=socket,host=localhost,port=2002;urp;', '--norestore'])
    time.sleep(10)

# Función para cerrar el servidor de LibreOffice
def stop_libreoffice_server():
    for proc in psutil.process_iter(['name', 'pid']):
        if proc.info['name'] == 'soffice.bin':
            print(f"Cerrando el servidor de LibreOffice con PID: {proc.info['pid']}")
            proc.terminate()  # Intentar terminar suavemente
            try:
                proc.wait(timeout=10)  # Esperar hasta 10 segundos para cerrar el proceso
                print("Servidor de LibreOffice cerrado correctamente.")
            except psutil.TimeoutExpired:
                print("El servidor de LibreOffice no se cerró en el tiempo esperado. Forzando cierre...")
                proc.kill()  # Forzar cierre si no responde
                print("Servidor de LibreOffice forzado a cerrar.")

# Función para obtener la ruta al intérprete de Python de LibreOffice
def get_libreoffice_python_path():
    try:
        soffice_executable = find_soffice_executable()
        libreoffice_base = os.path.dirname(os.path.dirname(soffice_executable))

        if platform.system() == 'Darwin':  # macOS
            python_libreoffice_path = os.path.join(libreoffice_base, 'Resources', 'python')
        elif platform.system() == 'Linux':
            python_libreoffice_path = os.path.join(libreoffice_base, 'program', 'python')
        elif platform.system() == 'Windows':
            python_libreoffice_path = os.path.join(libreoffice_base, 'program', 'python.exe')
        else:
            raise EnvironmentError("Sistema operativo no soportado.")

        if not os.path.isfile(python_libreoffice_path):
            raise FileNotFoundError(f"No se encontró el intérprete de Python de LibreOffice en '{python_libreoffice_path}'")

        return python_libreoffice_path
    except Exception as e:
        print(f"Error al obtener la ruta al intérprete de Python de LibreOffice: {e}")
        exit(1)

# Verificar si el servidor de LibreOffice está en ejecución
server_started = False

# Bandera para ejecutar solo una vez la extracción de datos
datos_extraidos = False

if not is_libreoffice_running() and not datos_extraidos:
    print("Iniciando el servidor de LibreOffice...")
    start_libreoffice_server()
    server_started = True

# Obtener la ruta absoluta del directorio actual
current_dir = os.path.dirname(os.path.abspath(__file__))

# Obtener la ruta al intérprete de Python de LibreOffice
python_libreoffice_path = get_libreoffice_python_path()

# Ruta relativa al script de extracción
extraccion_script_path = os.path.join(current_dir, 'scripts', 'multiple_consulta.py')

# Ejecutar el script de extracción de datos con el Python de LibreOffice solo una vez
if not datos_extraidos:
    try:
        print("Ejecutando el script de extracción de datos...")
        subprocess.run([python_libreoffice_path, extraccion_script_path], check=True)
        print("Extracción de datos completada.")
        datos_extraidos = True  # Marcar que los datos ya han sido extraídos
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar el script de extracción: {e}")
        if server_started:
            stop_libreoffice_server()
        exit(1)
    finally:
        # Asegurarse de que el servidor de LibreOffice se cierre si fue iniciado por el script
        if server_started:
            print("Cerrando el servidor de LibreOffice...")
            stop_libreoffice_server()

# Leer los CSV con encoding para evitar errores
rechazos_df = pd.read_csv('rechazos.csv', encoding='utf-8')
partes_abiertos_df = pd.read_csv('partes_abiertos_por_dia.csv', parse_dates=['FECHA APERTURA'], encoding='utf-8')
importe_df = pd.read_csv('importe_por_dia.csv', parse_dates=['FECHA CIERRE'], encoding='utf-8')

# Crear la aplicación de Dash
app = dash.Dash(__name__)

# Opciones de filtro (día, semana, mes)
frecuencia_opciones = [
    {'label': 'Día', 'value': 'D'},
    {'label': 'Semana', 'value': 'W'},
    {'label': 'Mes', 'value': 'M'}
]

# Diseño del Dashboard
app.layout = html.Div([
    html.H1("Cuadro de Mandos: Servicios y Facturación"),
    
    # Selector de rango de fechas
    dcc.DatePickerRange(
        id='rango-fechas',
        min_date_allowed=max(partes_abiertos_df['FECHA APERTURA'].min(), pd.to_datetime("2023-10-20")),
        max_date_allowed=pd.to_datetime("today"),
        start_date=pd.to_datetime("2024-09-01"),
        end_date=pd.to_datetime("today"),
        display_format='DD/MM/YY'  # Establecer el formato de visualización en dd/mm/aa

    ),
    
    # Filtro de frecuencia
    html.Label("Selecciona la Frecuencia:"),
    dcc.RadioItems(
        id='frecuencia-radio',
        options=frecuencia_opciones,
        value='W',  # Por defecto, mostrar por semana
        inline=True
    ),
    
    # Gráfico de Partes Abiertos
    dcc.Graph(id='grafico-partes-abiertos'),

    # Gráfico de Importe Facturado
    dcc.Graph(id='grafico-importe-facturado'),

    # Gráfico de Tarta (Rechazos)
    dcc.Graph(id='grafico-rechazos')
])


# Callback para actualizar gráficos según la frecuencia y el rango de fechas seleccionados
@app.callback(
    [Output('grafico-partes-abiertos', 'figure'),
     Output('grafico-importe-facturado', 'figure'),
     Output('grafico-rechazos', 'figure')],
    [Input('frecuencia-radio', 'value'),
     Input('rango-fechas', 'start_date'),
     Input('rango-fechas', 'end_date')]
)
def actualizar_graficos(frecuencia, start_date, end_date):
    # Convertir las fechas seleccionadas a formato datetime
    fecha_inicio = pd.to_datetime(start_date)
    fecha_fin = pd.to_datetime(end_date)

    # Definir un rango de fechas válidas
    fecha_inicio_valida = pd.to_datetime("2023-10-20")
    fecha_actual = pd.to_datetime("today")  # Tomar la fecha actual como límite superior

    # Filtrar fechas incorrectas en 'FECHA APERTURA' y 'FECHA CIERRE'
    partes_abiertos_df_filtrado = partes_abiertos_df[
    (partes_abiertos_df['FECHA APERTURA'] >= fecha_inicio_valida) &
    (partes_abiertos_df['FECHA APERTURA'] <= fecha_actual)
]

    importe_df_filtrado = importe_df[
    (importe_df['FECHA CIERRE'] >= fecha_inicio_valida) &
    (importe_df['FECHA CIERRE'] <= fecha_actual)
]

    # Filtrar los datos en función del rango de fechas seleccionado
    partes_abiertos_df_filtrado = partes_abiertos_df[
        (partes_abiertos_df['FECHA APERTURA'] >= fecha_inicio) &
        (partes_abiertos_df['FECHA APERTURA'] <= fecha_fin)
    ]

    importe_df_filtrado = importe_df[
        (importe_df['FECHA CIERRE'] >= fecha_inicio) &
        (importe_df['FECHA CIERRE'] <= fecha_fin)
    ]

    # Agrupar los datos según la frecuencia seleccionada (Día, Semana o Mes)
    partes_abiertos_agg = partes_abiertos_df_filtrado.resample(frecuencia, on='FECHA APERTURA').sum()
    importe_agg = importe_df_filtrado.resample(frecuencia, on='FECHA CIERRE').sum()

    # Calcular la media de los valores en el rango de fechas seleccionado
    media_partes_abiertos = partes_abiertos_agg["PARTES ABIERTOS"].mean()
    media_importe = importe_agg["IMPORTE TOTAL"].mean()

    # Gráfico de Partes Abiertos con línea de la media
    grafico_partes_abiertos = px.line(partes_abiertos_agg, x=partes_abiertos_agg.index, y="PARTES ABIERTOS",
                                      title=f"Partes Abiertos por {frecuencia}")
    grafico_partes_abiertos.add_hline(y=media_partes_abiertos, line_dash="dash", 
                                      annotation_text="Media", annotation_position="bottom right")

    # Gráfico de Importe Facturado con línea de la media
    grafico_importe_facturado = px.line(importe_agg, x=importe_agg.index, y="IMPORTE TOTAL",
                                        title=f"Importe Facturado por {frecuencia}")
    grafico_importe_facturado.add_hline(y=media_importe, line_dash="dash", 
                                        annotation_text="Media", annotation_position="bottom right")

    # Gráfico de Tarta (Rechazos vs No Rechazos)
    total_rechazos = rechazos_df["RECHAZO"].value_counts()
    grafico_rechazos = px.pie(names=['Rechazados', 'No Rechazados'], values=total_rechazos,
                              title="Rechazos vs No Rechazos")

    return grafico_partes_abiertos, grafico_importe_facturado, grafico_rechazos


# Iniciar la aplicación (sin debug para evitar reinicios)
if __name__ == '__main__':
    app.run_server(debug=False)  # Desactivar el modo debug para evitar la doble ejecución
