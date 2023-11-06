import pandas as pd
import pyautogui as pg
import webbrowser as web
import time

# LEEMOS EL EXCEL DONDE TENEMOS LOS DATOS

data = pd.read_excel("./Telefonos.xlsx")

# HACEMOS UN BUCLE PARA ITERAR POR CADA ITEM DEL EXCEL, i REPRESENTA CADA FILA

for i in range(len(data)):
    celular = data.loc[i, "Telefono"].astype(str) # EL NUMERO LO PASAMOS COMO STRING PARA PODER AGREGARLO AL LINK
    nombre = data.loc[i, "Nombre"]
    deuda = data.loc[i, "Deuda"].astype(str)

    mensaje = "Hola, " + nombre + " presentas una deuda de " + deuda + " ¡este es un mensaje automatico con Python, Saludos!"
    
    # EL WEB.OPEN VA A ABRIR EL EXPLORADOR DE INTERNET QUE TENGAS PREDETERMINADO, SE PUEDE PERSONALIZAR 
    
    #chrome_path = "RUTA A LA UBICACION DEL ARCHIVO chrome.exe"
    #web.get(chrome_path).open('https://web.whatsapp.com/send?phone=' + celular + '&text=' + mensaje)
    
    web.open('https://web.whatsapp.com/send?phone=' + celular + '&text=' + mensaje)

    # ENVIAR EL MENSAJE

    time.sleep(8)                     # TIEMPO DE ESPERA LUEGO DE ABIR LA PESTAÑA NUEVA
    pg.click(1150,985)                # CLICK EN LA BARRA DE MENSAJE (VA A DEPENDER DEL TAMAÑO DE LA PANTALLA)
    time.sleep(1.5)                   # TIEMPO DE ESPERA LUEGO DE CLICKEAR LA BARRA DE MENSAJE
    pg.press('enter')                 # PRESIONAR ENTER PARA ENVIAR EL MENSAJE

    # CERRAR LA PESTAÑA

    time.sleep(2)                     # TIEMPO DE ESPERA PRUDENTE PARA QUE SE ENVIE EL MENSAJE
    pg.hotkey('ctrl', 'w')            # CERRAR LA PESTAÑA
    time.sleep(2)                     # TIEMPO DE ESPERA HASTA VOLVER A INICIAR EL BUCLE