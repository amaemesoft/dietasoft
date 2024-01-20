import pandas as pd
import os
import tkinter as tk
from tkinter import simpledialog
import logging
from openpyxl import load_workbook
import requests
import io

# URL de la base de datos en GitHub
base_de_datos_url = "https://raw.githubusercontent.com/amaemesoft/dietasoft/main/basededatos.xlsx"

# Obtener la ubicación del script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Configuración del logging para registrar información importante y errores
logging.basicConfig(filename=os.path.join(script_dir, 'registro_errores.log'), level=logging.DEBUG, format='%(asctime)s %(levelname)s:%(message)s')

def consultar_base_de_datos(url):
    """Consulta la base de datos directamente desde la URL en GitHub."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        df = pd.read_excel(pd.ExcelFile(io.BytesIO(response.content)))
        return df
    except Exception as e:
        logging.error(f"Error al consultar la base de datos: {e}")
        return None

# Consultar la base de datos desde GitHub
base_de_datos = consultar_base_de_datos(base_de_datos_url)

def buscar_datos_trabajador(dni):
    """Busca los datos de un trabajador por DNI en un archivo Excel y devuelve un diccionario con los datos."""
    try:
        ruta_base_datos = os.path.join(script_dir, 'basededatos.xlsx')
        df = pd.read_excel(ruta_base_datos)
        datos_trabajador = df[df['DNI'] == dni]

        if datos_trabajador.empty:
            logging.warning(f"No se encontró el DNI {dni} en la base de datos.")
            return None
        else:
            return datos_trabajador.to_dict(orient='records')[0]
    except Exception as e:
        logging.error(f"Error al buscar datos del trabajador: {e}")
        return None

def abrir_documento():
    """Abre el documento Excel rellenado con los datos del trabajador."""
    try:
        root = tk.Tk()
        root.withdraw()
        dni = simpledialog.askstring("Input", "Introduce tu DNI:", parent=root)
        if dni:
            datos_trabajador = buscar_datos_trabajador(dni)
            if datos_trabajador is None:
                return

            ruta_modelo = seleccionar_modelo_dieta(datos_trabajador)
            ruta_documento_nuevo = rellenar_excel(ruta_modelo, datos_trabajador)
            if ruta_documento_nuevo is None:
                return

            os.startfile(ruta_documento_nuevo)
    except Exception as e:
        logging.error(f"Error al abrir el documento: {e}")

# Interfaz gráfica para pedir el DNI
abrir_documento()


def seleccionar_modelo_dieta(datos_trabajador):
    # Definir rutas relativas para los modelos de dieta
    modelos_dieta = {
        '01_PI_ACOGIDA ESTÁNDAR_Dieta': os.path.join(script_dir, '01_PI_ACOGIDA ESTÁNDAR_Dieta.xlsx'),
        '02_PI_ACOGIDA VULNERABLES_Dieta': os.path.join(script_dir, '02_PI_ACOGIDA VULNERABLES_Dieta.xlsx'),
        '03_PI_AUTONOMÍA_Dieta': os.path.join(script_dir, '03_PI_AUTONOMÍA_Dieta.xlsx'),
        '04_PI_SERVICIOS DE APOYO, INTERVENCIÓN Y ACOMPAÑAMIENTO_Dieta': os.path.join(script_dir, '04_PI_SERVICIOS DE APOYO, INTERVENCIÓN Y ACOMPAÑAMIENTO_Dieta.xlsx'),
        '05_PI_VALORACIÓN INICIAL Y DERIVACIÓN_Dieta': os.path.join(script_dir, '05_PI_VALORACIÓN INICIAL Y DERIVACIÓN_Dieta.xlsx')
    }
    return modelos_dieta.get(datos_trabajador['MODELO DIETAS'], os.path.join(script_dir, 'modelo_por_defecto.xlsx'))

def rellenar_excel(ruta_modelo, datos_trabajador):
    """Rellena un modelo de Excel con los datos del trabajador."""
    try:
        libro = load_workbook(ruta_modelo)
        hoja = libro.active
        
        # Mapeo de celdas en Excel a los datos correspondientes
        mapeo_celdas_datos = {
            'A9': datos_trabajador['TRABAJADOR/A'],  
            'A10': datos_trabajador['DNI'],           
            'A11': datos_trabajador['DIRECCIÓN CENTRO'],
            'A12': str(datos_trabajador['PROYECTO EN RRHH']),
            'A13': datos_trabajador['FASE EN RRHH'],
            'A14': str(datos_trabajador['ANALÍTICA DIETAS']),
            'A15': str(datos_trabajador['AREA DIETAS'])
        }

        # Definir ruta relativa para la nueva carpeta de documentos
        ruta_documentos = os.path.join(script_dir, 'documentos')

        # Crear la carpeta "documentos" si no existe
        os.makedirs(ruta_documentos, exist_ok=True)

        # Definir ruta relativa para el nuevo documento
        ruta_documento_nuevo = os.path.join(ruta_documentos, f"Dieta_{datos_trabajador['DNI']}.xlsx")

        # Rellenar las celdas con los datos
        for celda, dato in mapeo_celdas_datos.items():
            contenido_actual = hoja[celda].value or ""
            hoja[celda] = contenido_actual + " " + dato

        # Guardar el documento Excel en la ruta relativa
        libro.save(ruta_documento_nuevo)
        logging.info(f"Documento Excel guardado en: {ruta_documento_nuevo}")

        return ruta_documento_nuevo
    except Exception as e:
        logging.error(f"Error al rellenar el documento Excel: {e}")
        return None


def abrir_documento(dni):
    """Abre el documento Excel rellenado con los datos del trabajador."""
    try:
        ruta_base_datos = os.path.join(script_dir, 'basededatos.xlsx')
        
        datos_trabajador = buscar_datos_trabajador(dni, ruta_base_datos)
        if datos_trabajador is None:
            return

        ruta_modelo = seleccionar_modelo_dieta(datos_trabajador)
        ruta_documento_nuevo = rellenar_excel(ruta_modelo, datos_trabajador)
        if ruta_documento_nuevo is None:
            return

        os.startfile(ruta_documento_nuevo)
    except Exception as e:
        logging.error(f"Error al abrir el documento: {e}")

# Interfaz gráfica para pedir el DNI
root = tk.Tk()
root.withdraw()
dni = simpledialog.askstring("Input", "Introduce tu DNI:", parent=root)
if dni:
    abrir_documento(dni)