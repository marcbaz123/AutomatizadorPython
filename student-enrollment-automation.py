import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys

# Configurar la salida para manejar caracteres especiales
sys.stdout.reconfigure(encoding='utf-8')


# --- ARCHIVOS ---
excel_file = r"FICHERO EXCEL CON DATOS A CARGAR"
nome_hoja = "G2 SUB 1"
columna_nombres = "NOMBRE Y APELLIDO"
columna_catraca = "CI"

excel_referencia = r"FICHERO EXCEL CON NOMBRES EN FORMATO CORRECTO"
columna_backup = "nombre_apellido"
columna_catraca_referencia = "numero_catraca"

# --- Leer los Excel ---
df_principal = pd.read_excel(excel_file, sheet_name=nome_hoja)
df_referencia = pd.read_excel(excel_referencia)

# Limpiar encabezados de espacios de los DataFrames 
df_principal.columns = df_principal.columns.str.strip()
df_referencia.columns = df_referencia.columns.str.strip()

# --- Diccionario catraca → nombre en formato correcto ---
dict_referencia = pd.Series(
    df_referencia[columna_backup].values,
    index=df_referencia[columna_catraca_referencia]
).to_dict()

# --- Asegurar que los nombres están en el formato correcto ---
def garantir_nome(row):
    nome = str(row[columna_nombres]).strip()
    if ',' in nome:  # ya en formato "Apellido, Nombre"
        return nome
    else:
        catraca = row[columna_catraca]
        return dict_referencia.get(catraca, nome).strip()

df_principal[columna_nombres] = df_principal.apply(garantir_nome, axis=1)

# --- Preparar lista de alumnos ---
alumnos = df_principal[[columna_nombres, columna_catraca]].dropna()

# --- Login ---
url_login = "INTRODUCIR URL DEL LOGIN"
usuario = "INTRODUCIR USERNAME"
contrasena = "INTRODUCIR CONTRASEñA"

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--ignore-certificate-errors")
driver = webdriver.Chrome(options=options)

driver.get(url_login)
wait = WebDriverWait(driver, 5)

driver.find_element(By.NAME, "username").send_keys(usuario)
driver.find_element(By.NAME, "password").send_keys(contrasena)
driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)
time.sleep(3)

# --- Ir a planificación ---
url_planificacion = "URL DONDE SE CARGARAN LOS DATOS"
driver.get(url_planificacion)
time.sleep(3)

# --- RESULTADOS ---
resultados = []

for _, row in alumnos.iterrows():
    nombre = row[columna_nombres]
    catraca_planilla = row[columna_catraca]
    catraca_sys = None

    # buscar en el dataframe de referencia si esa catraca existe
    if catraca_planilla in df_referencia[columna_catraca_referencia].values:
        catraca_sys = df_referencia.loc[
            df_referencia[columna_catraca_referencia] == catraca_planilla,
            columna_catraca_referencia
        ].values[0]

    try:
        # --- Buscar nombre ---
        campo_busqueda = wait.until(EC.presence_of_element_located((By.ID, "buscado")))
        campo_busqueda.clear()
        campo_busqueda.send_keys(nombre)
        campo_busqueda.send_keys(Keys.RETURN)
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)

        # --- Revisar lista ---
        lista = wait.until(EC.presence_of_element_located((By.ID, "lista-inscriptos")))
        texto_lista = lista.text.strip()
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)

        if "No existen inscriptos o todos están matriculados" in texto_lista:
            print(f"⚠ {nombre}: ya matriculado o no existe.")
            resultados.append({
                "Catraca SYS": catraca_sys,
                "Catraca planilla": catraca_planilla,
                "Nombre": nombre,
                "Situacion": "NO INSCRIPTO"
            })
            time.sleep(3)
            continue

        # --- Clic en botón Agregar ---
        boton_agregar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.btn.btn-primary")))
        boton_agregar.click()

        resultados.append({
            "Catraca SYS": catraca_sys,
            "Catraca planilla": catraca_planilla,
            "Nombre": nombre,
            "Situacion": "INSCRIPTO"
        })

        time.sleep(3)

    except Exception as e:
        print(f"❌ Error con {nombre}: {e}")
        resultados.append({
            "Catraca SYS": catraca_sys,
            "Catraca planilla": catraca_planilla,
            "Nombre": nombre,
            "Situacion": "ERROR"
        })
        time.sleep(3)
        continue

# --- Guardar resultados ---
base, ext = os.path.splitext(excel_file)
novo_arquivo = f"{base}-resultado-{nome_hoja}{ext}"

df_resultados = pd.DataFrame(resultados)
df_resultados.to_excel(novo_arquivo, index=False)
print(f"✅ Resultados guardados en {novo_arquivo}")