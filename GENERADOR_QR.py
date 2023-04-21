import qrcode
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image

# conectar con la hoja de cálculo de Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\Gustavo\Python\Serviplus\Serviplus\key.json", scope)
cliente = gspread.authorize(creds)
hoja = cliente.open("FALTANTES").worksheet('MRC')

# obtener los datos de los empleados desde la hoja de cálculo
datos_de_empleados = hoja.get_all_records()

# generar un código QR para cada empleado y guardar como imagen
for empleado in datos_de_empleados:
    # obtener el identificador único del empleado
    identificador = empleado["NOMBRE"] # remplaza "numero_de_identificacion" por el nombre de la columna que contenga los identificadores de tus empleados
    agencia = empleado["AGENCIA"]
    regional = empleado["REGIONAL"]
    datos_empleado = f"{identificador}, {agencia}, {regional}"

    # crear el código QR y guardarlo como imagen
    codigo_qr = qrcode.make(datos_empleado)
    ruta_al_codigo_qr = r"C:\Users\Gustavo\Python\SCRIPT QR\CODIGOS QR\{}.png".format(identificador)
    codigo_qr.save(ruta_al_codigo_qr)

    # guardar el código QR en la hoja de cálculo
    fila_nueva = [identificador, agencia, regional, ruta_al_codigo_qr]
    hoja.append_row(fila_nueva)