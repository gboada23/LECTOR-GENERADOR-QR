import pyzbar.pyzbar as pyzbar
import cv2
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import datetime


# Definir las credenciales de la cuenta de servicio de Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('key.json', scope)
client = gspread.authorize(creds)

# Abrir la hoja de cálculo y seleccionar la hoja de trabajo
sheet = client.open('CONTROL ASISTENCIA').worksheet('ASISTENCIA')
# Definir la función para decodificar los códigos QR
def decode(im):
    # Decodificar el código QR utilizando la librería pyzbar
    decodificar = pyzbar.decode(im)
    return decodificar

# Definir la función para agregar la información del personal en la hoja de cálculo

def add_to_sheet(data):
    # Obtener la fecha y hora actual
    ahora = datetime.datetime.now()
    fecha = ahora.strftime("%d/%m/%Y")
    hora = ahora.strftime("%H:%M:%S")
    
    # Buscar la última fila que contenga el mismo nombre y fecha
    ultima_fila = len(sheet.get_all_values()) + 1
    nombre = data
    for i in range(ultima_fila-1, 0, -1):
        if sheet.cell(i, 1).value == nombre and sheet.cell(i, 2).value == fecha:
            ultima_fila = i
            break
    
    # Verificar si se debe agregar una nueva fila o actualizar una existente
    if ultima_fila == len(sheet.get_all_values()) + 1:
        # Agregar una nueva fila con la información del personal y la fecha y hora de entrada
        filas = [nombre, fecha, hora]
       
        sheet.append_row(filas)
        print("Entrando a trabajar:\n", data, " a las", hora)
        
    else:
        # Actualizar la fila existente con la hora de salida
        sheet.update_cell(ultima_fila, 4, hora)
        print("Saliendo de trabajar: \n", data, " a las", hora)
        
# Definir la función principal para tomar la asistencia
def main():
    # Iniciar la cámara y establecer la resolución
    cap = cv2.VideoCapture(0)
    cap.set(3, 640)
    cap.set(4, 480)
    
    # Definir el tiempo mínimo entre detecciones del mismo código QR (en segundos)
    tiempo_de_detencion = 5
    ultimo_tiempo_detencion = 0
    while True:
        # Leer la imagen de la cámara
        ret, frame = cap.read()
        
        # Decodificar el código QR de la imagen
        decodificar = decode(frame)
        cv2.putText(frame, "Coloca el Codigo QR" ,(155, 80), cv2.FONT_HERSHEY_SIMPLEX, 1,(0, 0, 255), 2) 
        cv2.rectangle(frame, (170, 100), (470, 400), (0, 255, 0), 2) 
        # Si se ha detectado un código QR, agregar la información del personal a la hoja de cálculo
        if len(decodificar) > 0:
            data = decodificar[0].data.decode("utf-8")
            if time.time() - ultimo_tiempo_detencion >= tiempo_de_detencion:
                add_to_sheet(data)
                ultimo_tiempo_detencion = time.time()
                x, y, w, h = decodificar[0].rect
                cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)
                # Obtener el nombre de la persona y mostrarlo encima del rectángulo delimitador
                nombre = decodificar[0].data.decode("utf-8")
                cv2.putText(frame, nombre, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 0, 0), 2)            
            
            else:
                x, y, w, h = decodificar[0].rect
                cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)
                # Obtener el nombre de la persona y mostrarlo encima del rectángulo delimitador
                nombre = decodificar[0].data.decode("utf-8")
                cv2.putText(frame, nombre, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 0, 0), 2) 
        # Mostrar la imagen en una ventana
        cv2.imshow('DETECTOR DE QR', frame)
        
        # Si se presiona la tecla 'q', salir del bucle
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
        
if __name__ == '__main__':
    main()
    