Student Enrollment Automation

Automatización en Python para la carga masiva de alumnos desde archivos Excel a un sistema interno. El script realiza validación de datos, estandarización de nombres y registra los resultados en un archivo Excel final.

🔹 Características

Lectura de datos desde Excel (.xlsx).

Normalización de nombres usando un archivo de referencia.

Automatización web con Selenium: login, búsqueda y carga de alumnos.

Registro de alumnos inscritos, no inscritos o con errores en un Excel final.

Compatible con Windows y Linux (requiere Chrome y ChromeDriver).

🔹 Requisitos

Python 3.8+

Librerías:

pip install pandas selenium openpyxl


Google Chrome instalado.

ChromeDriver correspondiente a tu versión de Chrome (descargar aquí
).

🔹 Configuración

Renombrar tu archivo Excel de datos y la hoja a usar:

```excel_file = r"FICHERO_EXCEL_CON_DATOS.xlsx"
nome_hoja = "NOMBRE_HOJA"
columna_nombres = "NOMBRE Y APELLIDO"
columna_catraca = "CI"


Configurar archivo de referencia con nombres correctos:

excel_referencia = r"FICHERO_EXCEL_DE_REFERENCIA.xlsx"
columna_backup = "nombre_apellido"
columna_catraca_referencia = "numero_catraca"


Configurar login del sistema:

url_login = "URL_DEL_LOGIN"
usuario = "TU_USUARIO"
contrasena = "TU_CONTRASEÑA"
url_planificacion = "URL_DE_CARGA_DE_DATOS"

python student-enrollment-automation.py


El script hará lo siguiente:

Leer y limpiar los datos de Excel.

Normalizar nombres usando el archivo de referencia.

Abrir Chrome y hacer login en el sistema interno.

Buscar cada alumno y cargarlo automáticamente.

Guardar los resultados en un Excel llamado:

<nombre_archivo>-resultado-<nombre_hoja>.xlsx


Mostrando si cada alumno fue inscrito, no inscrito o tuvo error.

🔹 Estructura de archivos recomendada
AutomatizadorPython/
│
├─ student-enrollment-automation.py
├─ datos.xlsx               # Archivo principal con alumnos
├─ referencia.xlsx          # Archivo con nombres correctos
└─ README.md

🔹 Notas importantes

Asegúrate de que ChromeDriver esté en el PATH o en la misma carpeta del script.

Verifica que los nombres de columnas coincidan exactamente con tus archivos Excel.

Se recomienda ejecutar con una cuenta de prueba antes de usar en producción.
