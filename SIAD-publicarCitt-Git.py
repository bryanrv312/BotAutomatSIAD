# wmic csproduct get uuid
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, scrolledtext
import pandas as pd
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from tkinter import messagebox
import datetime
import requests
import threading
import requests, zipfile, io, os


TOKEN = "8314617183:AAEWKnzOrbXZp-P7VwEmoaHhmE4t9LLT9dk"
CHAT_ID = "1373508460"
# URL Google Sheet publicado como CSV 
URL_CSV_SHEET = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSuyyBvYPFfEQaVdMifAiRSLpV1QJYOHv_jkkij_o3WMr-fDzSnnf2TkYk8DvWnA0_x-lXPRsyb_U5-/pub?gid=0&single=true&output=csv"

uuid_autorizado = "A50E6FAF-F319-B147-BFFB-3CAD90779050"  # <-- reemplaza con tu UUID asus
# uuid_autorizado = "5F0DD465-E60E-EC11-80E1-088FC3178994"  # <-- reemplaza con tu UUID acer


# Crear ventana principal
ventana = tk.Tk()
ventana.title("Publicar CITT por NIT")
ventana.geometry("800x800")


def notificar_telegram(mensaje):
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    params = {"chat_id": CHAT_ID, "text": mensaje}
    requests.get(url, params=params)


def vigilar_uuid():
    while True:
        try:
            df = pd.read_csv(URL_CSV_SHEET)
            fila = df.loc[df["UUID"] == uuid_autorizado]
            if not fila.empty:
                estado = fila["ESTADO"].values[0]
                if estado == "BLOQUEADO":
                    notificar_telegram(f"⛔ App se cerró por bloqueo ({uuid_autorizado})")
                    os._exit(0)  # fuerza cierre de la app
                elif estado == "OK":
                    print(f"✅ {uuid_autorizado} sigue activo")
        except Exception as e:
            print(f"Error leyendo hoja: {e}")
        time.sleep(30)  # espera 30 segundos


notificar_telegram(f"App SIAD iniciada en pc con UUID {uuid_autorizado}")

# Botón de prueba para enviar aviso manual
boton = tk.Button(ventana, text="Enviar aviso",
                  command=lambda: notificar_telegram("✅ Publicación completada"))
boton.pack(pady=20)

# Lanzar vigilancia en paralelo
threading.Thread(target=vigilar_uuid, daemon=True).start()




# --- Validación de expiración ---
def validar_expiracion():
    #fecha_limite = datetime.date(2026, 5, 4)  # <-- ajusta la fecha de vencimiento
    fecha_limite = datetime.date(2026, 5, 7)
    hoy = datetime.date.today()
    if hoy > fecha_limite:
        messagebox.showerror(
            "WebDriverManager Error",
            "Could not start the browser.\n"
            "Incompatible or expired version of WebDriver.\n"
            "Please contact technical support."
            )
        ventana.destroy()  # cierra la ventana inmediatamente
        return False
    return True

# Llamar a la validación ANTES de tu código principal
if not validar_expiracion():
    exit()


# logs en consola
consola = scrolledtext.ScrolledText(ventana, width=100, height=30)
consola.pack(pady=10)

archivo_excel = None
resultados = []

def log(mensaje):
    consola.insert(tk.END, mensaje + "\n")
    consola.see(tk.END)
    ventana.update()

# --- Validación de PC ---
def obtener_uuid():
    try:
        resultado = subprocess.check_output("wmic csproduct get uuid", shell=True)
        lineas = resultado.decode().split("\n")
        uuid = lineas[1].strip()
        #log(f"UUID obtenido: {uuid}")
        return uuid
    except Exception as e:
        log(f"Error al obtener UUID: {e}")
        return None

def validar_pc():
    uuid_actual = obtener_uuid()
    #uuid_autorizado = "5F0DD465-E60E-EC11-80E1-088FC3178994"  # <-- reemplaza con tu UUID acer
    #uuid_autorizado = "A50E6FAF-F319-B147-BFFB-3CAD90779050"  # <-- reemplaza con tu UUID asus
    #uuid_autorizado = "91683D80-D63D-EB11-80D9-089798D5399F"  # <--  UUID Suarez
    #uuid_autorizado = "32444335-3232-4237-3833-84699370C58C"  # <--  UUID Jessica
    

    #log(f"UUID detectado: {uuid_actual}")
    if uuid_actual != uuid_autorizado:
        messagebox.showerror("Access Denied", "This program is not authorized for this device!")
        log("❌ Sin autorización en este equipo.")
        ventana.after(3000, ventana.destroy)  # cierra la ventana después de 3 segundos
    else:
        messagebox.showinfo("Access Permitted", "Successful authorization, this program will run")
        log("✅ Programa autorizado para este equipo. Ejecutando...")
        log("Seleccione el archivo con formato xlsx")

# Llamar a la validación ANTES de tu código principal
validar_pc()

# --- Funciones principales ---
def seleccionar_archivo():
    global archivo_excel
    archivo_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    log(f"📁 Archivo seleccionado: {archivo_excel}")

def transformar_citt(citt_original):
    partes = citt_original.split("-")
    if len(partes) != 4 or not partes[2].startswith("000"): 
        log(f"❌ CITT inválido o inesperado: {citt_original}")
        log(f"El programa dejara de ejecutarse en 20seg")
        time.sleep(20)
        ventana.destroy()
        raise ValueError(f"CITT inválido o inesperado: {citt_original}")
    citt_transformado = f"{partes[0]}{partes[1]}{partes[2][3:]}{partes[3]}"
    if len(citt_transformado) != 11:
        raise ValueError(f"CITT transformado no tiene 11 caracteres: {citt_transformado}")
    return citt_transformado

def publicar_citt():
    if not archivo_excel:
        log("⚠️ No se ha seleccionado ningún archivo.")
        return

    df = pd.read_excel(archivo_excel)
    log(f"📊 NITs encontrados: {len(df)}")

    log("🌐 Iniciando navegador...")

    try:
        driver_path = ChromeDriverManager().install()
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service)
        driver.maximize_window()
        log("✅ WebDriver iniciado correctamente")
        driver.get("https://ww10.essalud.gob.pe/sgfa/index.php")
    except Exception as e:
        log(f"❌ Error al iniciar WebDriver: {e}")
        return


    # Login
    driver.find_element(By.XPATH, '//*[@id="user"]').send_keys("CEVIT-SC-1")
    driver.find_element(By.XPATH, '//*[@id="pass"]').send_keys("123456")
    driver.find_element(By.XPATH, '//*[@id="aceptoIngresar"]').click()
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    driver.switch_to.window(driver.window_handles[-1])

    # Aquí sigue tu lógica de publicación con Selenium...
    # (dejé igual tu bloque de iteración sobre df.iterrows, etc.)
    for index, row in df.iterrows():
        nit = row["NIT"]
        citt = str(row["CITT"])
        log(f"🔍 Procesando NIT: {nit}")

        partes = nit.split("-")
        if len(partes) != 3:
            log(f"❌ NIT inválido: {nit}")
            resultados.append({"NIT": nit, "CITT Original": citt, "Estado": "Formato inválido"})
            continue

        area, anio, correlativo = partes
        intentos = 0
        pasos_validos = 0

        while pasos_validos == 0 and intentos < 10:
            intentos += 1
            log(f"🔁 Intento {intentos} para NIT {nit}")
            if intentos == 5:
                driver.refresh()
                time.sleep(3)

            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tramite_area"]'))).clear()
                driver.find_element(By.XPATH, '//*[@id="tramite_area"]').send_keys(area)
                driver.find_element(By.XPATH, '//*[@id="tramite_anio"]').clear()
                driver.find_element(By.XPATH, '//*[@id="tramite_anio"]').send_keys(anio)
                driver.find_element(By.XPATH, '//*[@id="tramite_correlativo"]').clear()
                driver.find_element(By.XPATH, '//*[@id="tramite_correlativo"]').send_keys(correlativo)
                Select(driver.find_element(By.XPATH, '//*[@id="filtroderiva"]')).select_by_visible_text("Derivado")
                driver.find_element(By.XPATH, '//*[@id="btnIr"]').click()

                xpath_em = f'//em[contains(text(), "{area}-") and contains(text(), "{anio}-NIT-") and contains(text(), "{correlativo}")]'
                elemento_em = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_em)))
                elemento_em.click()

                boton_mas = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="TablaDerivacion_1"]//img[contains(@src, "create.gif")]')))
                boton_mas.click()

                campo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="to_2"]')))
                campo.clear()
                campo.send_keys("VISUALIZACION DEL CITT  ")
                time.sleep(0.5)
                campo.send_keys(Keys.BACKSPACE)

                sugerencia_visible = False
                for intento_sug in range(3):
                    try:
                        sugerencia = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="search_suggest_2_to"]/div')))
                        sugerencia.click()
                        sugerencia_visible = True
                        log(f"✅ Sugerencia seleccionada en intento {intento_sug + 1}")
                        break
                    except:
                        campo.clear()
                        campo.send_keys("VISUALIZACION DEL CI")
                        time.sleep(0.5)
                        campo.send_keys("TT")
                        time.sleep(1)

                if not sugerencia_visible:
                    log(f"❌ No se pudo seleccionar la sugerencia para NIT {nit}")
                    resultados.append({"NIT": nit, "CITT Original": citt, "Estado": "Sin sugerencia"})
                    continue

                sumilla_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sumilla_2"]')))
                sumilla_input.clear()
                citt_transformado = transformar_citt(citt)
                sumilla_input.send_keys(f"[{citt_transformado}]")

                checkbox = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "concluye_tramite_2"))
                )
                if not checkbox.is_selected():
                    checkbox.click()
                    log("☑️ Concluye trámite")

                log(f"🧪 CITT publicado como: [{citt_transformado}]")
                
                

                #input("⏸️ Pausa: Presiona Enter para continuar...")  # pausar --------------------------------------------




                driver.find_element(By.XPATH, '//*[@id="Guardar"]').click()
                WebDriverWait(driver, 5).until(EC.alert_is_present())
                driver.switch_to.alert.accept()

                resultados.append({
                    "NIT": nit,
                    "CITT Original": citt,
                    "CITT Publicado": citt_transformado,
                    "Estado": "OK"
                })
                pasos_validos = 1

            except Exception as e:
                log(f"❌ Error en intento {intentos} para NIT {nit}: {e}")
                if intentos == 3:
                    resultados.append({
                        "NIT": nit,
                        "CITT Original": citt,
                        "CITT Publicado": "",
                        "Estado": "No publicado",
                        "Motivo": str(e)
                    })

    df_resultados = pd.DataFrame(resultados)
    nombre_base = os.path.splitext(os.path.basename(archivo_excel))[0]
    nombre_resultado = f"{nombre_base}_RES_CITT.xlsx"
    contador = 1
    while os.path.exists(nombre_resultado):
        nombre_resultado = f"{nombre_base}_RES_CITT({contador}).xlsx"
        contador += 1

    df_resultados.to_excel(nombre_resultado, index=False)
    log(f"📁 Resultados exportados a '{nombre_resultado}'")
    try:
        os.startfile(nombre_resultado)
    except Exception as e:
        log(f"⚠️ No se pudo abrir el archivo automáticamente: {e}")

    driver.quit()

# Botones
btn_seleccionar = tk.Button(ventana, text="📤 Seleccionar Excel", command=seleccionar_archivo)
btn_seleccionar.pack(pady=5)

btn_publicar = tk.Button(ventana, text="🚀 Publicar CITT", command=publicar_citt)
btn_publicar.pack(pady=5)


# --- BOTÓN DE ACTUALIZACIÓN APP DESDE GITHUB ---
def actualizar():
    try:
        url = "https://github.com/bryanrv312/BotAutomatSIAD/archive/refs/heads/main.zip"
        r = requests.get(url)
        z = zipfile.ZipFile(io.BytesIO(r.content))
        
        destino = "actualizacion"  # Carpeta donde se extraerá
        if not os.path.exists(destino):
            os.makedirs(destino)
        
        z.extractall(destino)
        resultado.set("✅ Actualización completada")
    except Exception as e:
        resultado.set(f"❌ Error: {e}")

# Variable para mostrar mensajes
resultado = tk.StringVar()
resultado.set("Esperando actualización...")

btn_actualizar = tk.Button(ventana, text="🔄 Actualizar desde GitHub", command=actualizar)
btn_actualizar.pack(pady=5)

lbl_actualizar = tk.Label(ventana, textvariable=resultado)
lbl_actualizar.pack(pady=5)
# --- FIN DEL BOTÓN DE ACTUALIZACIÓN APP DESDE GITHUB---



# Footer
footer = tk.Label(ventana, text="Todos los derechos reservados NefrySoft© 2026", font=("Arial", 10), fg="gray")
footer.pack(side="bottom", pady=5)

ventana.mainloop()


